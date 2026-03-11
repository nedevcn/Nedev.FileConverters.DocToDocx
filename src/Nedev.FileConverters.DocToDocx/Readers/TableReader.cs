using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Reference equality comparer for HashSet since ReferenceEqualityComparer is not available in netstandard2.1
/// </summary>
internal class ReferenceEqualityComparer<T> : IEqualityComparer<T> where T : class
{
    public bool Equals(T? x, T? y) => ReferenceEquals(x, y);
    public int GetHashCode(T obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
}

/// <summary>
/// A state context for an actively parsed table level.
/// </summary>
internal class TableParseContext
{
    public int Level { get; set; }
    public TableModel Table { get; set; } = new();
    public int RowIndex { get; set; }
    public List<TableCellModel> CellsInCurrentRow { get; } = new();
    public List<ParagraphModel> CurrentCellParagraphs { get; } = new();
    public TapBase? CurrentRowTap { get; set; }
    public List<TapBase?> RowTaps { get; } = new();
    public int LastTableParagraphIndex { get; set; }
}

/// <summary>
/// Extracts tables (including nested tables) by interpreting paragraph nesting levels
/// and cell boundaries (\x07).
/// </summary>
public class TableReader
{
    private const int MaxNestingDepth = 20;
    private const int MaxRowsPerTable = 10000;
    private static readonly string _debugLogPath = Path.Combine(Path.GetTempPath(), "table_reader_debug.log");
    private sealed class RecoveredCell
    {
        public string Text { get; set; } = string.Empty;
        public ParagraphModel? SourceParagraph { get; set; }
        public List<RunModel> SourceRuns { get; } = new();
    }

    private readonly BinaryReader _wordDocReader;
    private readonly BinaryReader _tableReader;
    private readonly FibReader _fib;
    private readonly FkpParser _fkpParser;

    public TableReader(BinaryReader wordDocReader, BinaryReader tableReader, FibReader fib, FkpParser fkpParser)
    {
        _wordDocReader = wordDocReader;
        _tableReader = tableReader;
        _fib = fib;
        _fkpParser = fkpParser;
    }

    public void ParseTables(DocumentModel document)
    {
        try { File.Delete(_debugLogPath); }
        catch (Exception ex) { Logger.Debug($"Failed to delete table-reader debug log '{_debugLogPath}': {ex.Message}"); }
        Log("ParseTables START");
        var topLevelTables = new List<TableModel>();
        // Track all tables (including nested) to prevent duplicate additions
        // Use object.GetHashCode for reference equality since ReferenceEqualityComparer is not available in netstandard2.1
        var allTables = new HashSet<TableModel>(new ReferenceEqualityComparer<TableModel>());
        var stack = new List<TableContext>(); // index 0 = level 1, index 1 = level 2, etc.

        foreach (var para in document.Paragraphs.OrderBy(p => p.Index))
        {
            int level = para.NestingLevel <= 0 ? 0 : Math.Max(1, para.NestingLevel - 1);
            if (level > MaxNestingDepth) level = MaxNestingDepth;
            if (para.Index % 10 == 0 || level > 0)
            {
                Log($"Paragraph {para.Index} nesting={para.NestingLevel} adjustedLevel={level}");
            }
            if (para.Type != ParagraphType.TableCell && level == 0)
            {
                // Not in a table
                level = 0;
            }

            // If we dropped levels, finalize the deeper ones
            while (stack.Count > level)
            {
                var popped = stack.Last();
                stack.RemoveAt(stack.Count - 1);
                Log($"Popped TableContext level={popped.Level} at para={para.Index}");
                
                FlushCurrentCell(popped);
                FlushCurrentRow(popped);
                FinalizeTable(popped);

                // If the popped table actually has content, attach it
                // Note: We check allTables to prevent duplicate additions, but nested tables should always be attached to parent
                bool isNestedTable = stack.Count > 0;
                if (popped.Table.Rows.Count > 0 && (isNestedTable || !allTables.Contains(popped.Table)))
                {
                    if (stack.Count > 0)
                    {
                        var parent = stack.Last();
                        // record parent index for easier navigation later
                        popped.Table.ParentTableIndex = parent.Table.Index;
                        Log($"Set ParentTableIndex={parent.Table.Index} for table at para={popped.Table.StartParagraphIndex}");
                        
                        // Attach to current cell of parent
                        var nestedPara = new ParagraphModel
                        {
                            Type = ParagraphType.NestedTable,
                            NestedTable = popped.Table,
                            NestingLevel = parent.Level
                        };
                        var firstText = popped.Table.Rows.Count > 0 && popped.Table.Rows[0].Cells.Count > 0 && popped.Table.Rows[0].Cells[0].Paragraphs.Count > 0 
                            ? popped.Table.Rows[0].Cells[0].Paragraphs[0].Text : "(empty)";
                        Log($"Created NestedTable paragraph, nested table has {popped.Table.Rows.Count} rows, first cell text: '{firstText}'");
                        parent.CurrentCellParagraphs.Add(nestedPara);
                        allTables.Add(popped.Table);
                    }
                    else
                    {
                        // top-level table has no parent - ensure ParentTableIndex is null
                        popped.Table.ParentTableIndex = null;
                        var tableId = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(popped.Table);
                        Log($"Adding to topLevelTables (1st): table id={tableId} index={popped.Table.Index} startPara={popped.Table.StartParagraphIndex} ParentTableIndex={popped.Table.ParentTableIndex}");
                        topLevelTables.Add(popped.Table);
                        allTables.Add(popped.Table);
                    }
                }
            }

            if (level == 0) continue;

            // Ensure we have active contexts up to the current level
            while (stack.Count < level)
            {
                // Protect against pathological documents that claim absurd nesting levels.
                if (stack.Count >= MaxNestingDepth)
                {
                    break;
                }

                var newCtx = new TableContext
                {
                    Level = stack.Count + 1,
                    Table = new TableModel { Index = topLevelTables.Count + stack.Count, StartParagraphIndex = para.Index }
                };
                Log($"Pushing new TableContext level={newCtx.Level} startPara={para.Index}");
                stack.Add(newCtx);
            }

            foreach (var context in stack)
            {
                context.LastTableParagraphIndex = para.Index;
            }

            var activeCtx = stack.Last();

            // Extract TAP properties for table alignment and cell widths
            TapBase? tapForParagraph = null;
            var firstRun = para.Runs.FirstOrDefault();
            if (firstRun != null)
            {
                var pap = _fkpParser.GetPapAtCp(firstRun.CharacterPosition);
                tapForParagraph = pap?.Tap;

                if (tapForParagraph != null)
                {
                    // Merge TAP properties from all paragraphs, not just the first one
                    if (activeCtx.Table.Properties == null)
                    {
                        activeCtx.Table.Properties = new TableProperties();
                    }
                    
                    var props = activeCtx.Table.Properties;
                    
                    // Only set values if not already set (first valid value wins)
                    if (props.Alignment == TableAlignment.Left && tapForParagraph.Justification != 0)
                    {
                        props.Alignment = tapForParagraph.Justification switch
                        {
                            1 => TableAlignment.Center,
                            2 => TableAlignment.Right,
                            _ => TableAlignment.Left
                        };
                    }
                    
                    if (props.CellSpacing == 0)
                    {
                        props.CellSpacing = tapForParagraph.CellSpacing != 0
                            ? tapForParagraph.CellSpacing
                            : (tapForParagraph.GapHalf != 0 ? tapForParagraph.GapHalf * 2 : 0);
                    }
                    
                    if (props.Indent == 0 && tapForParagraph.IndentLeft != 0)
                    {
                        props.Indent = tapForParagraph.IndentLeft;
                    }
                    
                    if (props.PreferredWidth == 0 && tapForParagraph.TableWidth != 0)
                    {
                        props.PreferredWidth = tapForParagraph.TableWidth;
                    }
                    
                    // Borders - only set if not already present
                    props.BorderTop ??= tapForParagraph.BorderTop;
                    props.BorderBottom ??= tapForParagraph.BorderBottom;
                    props.BorderLeft ??= tapForParagraph.BorderLeft;
                    props.BorderRight ??= tapForParagraph.BorderRight;
                    props.BorderInsideH ??= tapForParagraph.BorderInsideH;
                    props.BorderInsideV ??= tapForParagraph.BorderInsideV;
                    props.Shading ??= tapForParagraph.Shading;
                }

                if (activeCtx.CurrentRowTap == null && tapForParagraph != null)
                {
                    activeCtx.CurrentRowTap = tapForParagraph;
                }
            }

            if (TryConsumeCompactTableParagraph(para, activeCtx, tapForParagraph))
            {
                continue;
            }

            // Cell boundary detection. Row end is a cell with ONLY \x07.
            string textContent = string.Join("", para.Runs.Select(r => r.Text));
            bool isRowEnd = string.IsNullOrWhiteSpace(textContent.Replace("\x07", ""));

            if (isRowEnd)
            {
                FlushCurrentCell(activeCtx);
                FlushCurrentRow(activeCtx);
            }
            else
            {
                // Accumulate paragraph into current cell.
                var cellParagraph = new ParagraphModel
                {
                    Index = para.Index,
                    Type = ParagraphType.Normal,
                    Properties = para.Properties,
                    NestingLevel = para.NestingLevel,
                    Runs = para.Runs.Select(r =>
                    {
                        // Clone the run but strip cell-end markers from text.
                        // Reuse the original Properties object to preserve ALL
                        // formatting (bold, italic, CS props, RGB, highlight, etc.)
                        var cloned = new RunModel
                        {
                            Text = r.Text?.Replace("\x07", "") ?? string.Empty,
                            IsPicture = r.IsPicture,
                            ImageIndex = r.ImageIndex,
                            FcPic = r.FcPic,
                            CharacterPosition = r.CharacterPosition,
                            CharacterLength = r.CharacterLength,
                            IsField = r.IsField,
                            FieldCode = r.FieldCode,
                            IsHyperlink = r.IsHyperlink,
                            HyperlinkUrl = r.HyperlinkUrl,
                            HyperlinkBookmark = r.HyperlinkBookmark,
                            HyperlinkRelationshipId = r.HyperlinkRelationshipId,
                            IsBookmark = r.IsBookmark,
                            IsBookmarkStart = r.IsBookmarkStart,
                            BookmarkName = r.BookmarkName,
                            IsOle = r.IsOle,
                            OleObjectId = r.OleObjectId,
                            OleProgId = r.OleProgId,
                            ImageRelationshipId = r.ImageRelationshipId,
                            Properties = r.Properties
                        };
                        return cloned;
                    }).ToList()
                };

                cellParagraph.Runs.RemoveAll(r => string.IsNullOrEmpty(r.Text) && !r.IsPicture);
                activeCtx.CurrentCellParagraphs.Add(cellParagraph);

                if (textContent.Contains('\x07'))
                {
                    // This paragraph ended the cell.
                    FlushCurrentCell(activeCtx);
                }
            }
        }

        while (stack.Count > 0)
        {
            var popped = stack.Last();
            stack.RemoveAt(stack.Count - 1);
            FlushCurrentCell(popped);
            FlushCurrentRow(popped);
            FinalizeTable(popped);
            // Only process tables that haven't been added yet
            // and only add to top-level (nested tables are already attached to parents in the main loop)
            if (popped.Table.Rows.Count > 0 && !allTables.Contains(popped.Table))
            {
                // This is a top-level table (no parent in stack)
                popped.Table.ParentTableIndex = null;
                Log($"Adding to topLevelTables (final): table index={popped.Table.Index} startPara={popped.Table.StartParagraphIndex} ParentTableIndex={popped.Table.ParentTableIndex}");
                topLevelTables.Add(popped.Table);
                allTables.Add(popped.Table);
            }
        }

        Log($"TopLevelTables={topLevelTables.Count}");
        document.Tables = topLevelTables;
        Log($"Assigned document.Tables count={document.Tables?.Count ?? 0}");

        if (!ContainsNestedTables(document.Tables ?? Enumerable.Empty<TableModel>()))
        {
            Log("ContainsNestedTables: false");
            if (NeedsFlatTableRecovery(document))
            {
                Log("RecoverFlatTables START");
                RecoverFlatTables(document);
                Log("RecoverFlatTables DONE");
            }

            if (!ContainsNestedTables(document.Tables ?? Enumerable.Empty<TableModel>()))
            {
                RecoverNestedTableSections(document);
            }
        }
        else
        {
            Log("ContainsNestedTables: true");
        }

        ApplyFallbackHeaderShading(document);

        Log("ParseTables END");
    }

    private static void ApplyFallbackHeaderShading(DocumentModel document)
    {
        var fallbackShading = CreateFallbackHeaderShading(document.Theme);
        if (fallbackShading == null)
            return;

        foreach (var table in document.Tables)
        {
            ApplyFallbackHeaderShading(table, fallbackShading);
        }
    }

    private static void ApplyFallbackHeaderShading(TableModel table, ShadingInfo fallbackShading)
    {
        foreach (var row in table.Rows)
        {
            if (!NeedsFallbackHeaderShading(row))
                continue;

            foreach (var cell in row.Cells)
            {
                cell.Properties ??= new TableCellProperties();
                cell.Properties.Shading ??= CloneShading(fallbackShading);
            }
        }

        foreach (var row in table.Rows)
        {
            foreach (var cell in row.Cells)
            {
                foreach (var paragraph in cell.Paragraphs)
                {
                    if (paragraph.NestedTable != null)
                    {
                        ApplyFallbackHeaderShading(paragraph.NestedTable, fallbackShading);
                    }
                }
            }
        }
    }

    private static bool NeedsFallbackHeaderShading(TableRowModel row)
    {
        if (row.Cells.Count == 0)
            return false;

        if (row.Cells.Any(CellHasExplicitShading))
            return false;

        int visibleRunCount = 0;
        int whiteRunCount = 0;

        foreach (var run in row.Cells
            .SelectMany(cell => cell.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs)
            .Where(run => !string.IsNullOrWhiteSpace(run.Text)))
        {
            visibleRunCount++;
            if (IsWhiteRun(run))
            {
                whiteRunCount++;
            }
        }

        if (visibleRunCount == 0)
            return false;

        return whiteRunCount == visibleRunCount;
    }

    private static bool CellHasExplicitShading(TableCellModel cell)
    {
        if (cell.Properties?.Shading != null)
            return true;

        return cell.Paragraphs.Any(paragraph => paragraph.Properties?.Shading != null);
    }

    private static bool IsWhiteRun(RunModel run)
    {
        var props = run.Properties;
        if (props == null)
            return false;

        if (props.HasRgbColor)
            return props.RgbColor == 0xFFFFFFu;

        return props.Color == 8;
    }

    private static ShadingInfo? CreateFallbackHeaderShading(ThemeModel theme)
    {
        // Prefer the document theme accent so missing table-style fills still match
        // the rest of the document palette when the binary source only preserves
        // inverse-video text.
        int backgroundColor = theme.ColorMap.ContainsKey("accent1")
            ? 0x01000004
            : 0x00BD814F;

        return new ShadingInfo
        {
            Pattern = ShadingPattern.Clear,
            PatternVal = "clear",
            BackgroundColor = backgroundColor
        };
    }

    private static ShadingInfo CloneShading(ShadingInfo shading)
    {
        return new ShadingInfo
        {
            Pattern = shading.Pattern,
            PatternVal = shading.PatternVal,
            ForegroundColor = shading.ForegroundColor,
            BackgroundColor = shading.BackgroundColor
        };
    }

    private static bool ContainsNestedTables(IEnumerable<TableModel> tables)
    {
        foreach (var table in tables)
        {
            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var paragraph in cell.Paragraphs)
                    {
                        if (paragraph.Type != ParagraphType.NestedTable || paragraph.NestedTable == null)
                            continue;

                        return true;
                    }
                }
            }
        }

        return false;
    }

    private bool NeedsFlatTableRecovery(DocumentModel document)
    {
        var flatTableParagraphs = document.Paragraphs
            .Where(LooksLikeFlatTableParagraph)
            .ToList();

        if (flatTableParagraphs.Count == 0)
            return false;

        int hintedColumnCount = EstimateColumnCount(flatTableParagraphs);
        if (hintedColumnCount < 2)
            return false;

        if (document.Tables.Count == 0)
            return true;

        return document.Tables.All(table =>
            table.RowCount <= 1 ||
            table.ColumnCount < hintedColumnCount ||
            table.Rows.All(row => row.Cells.Count < hintedColumnCount));
    }

    private void RecoverFlatTables(DocumentModel document)
    {
        Log("RecoverFlatTables(method) START");
        var recoveredTables = new List<TableModel>();

        for (int i = 0; i < document.Paragraphs.Count; i++)
        {
            if (!LooksLikeFlatTableParagraph(document.Paragraphs[i]))
                continue;

            int start = i;
            while (i + 1 < document.Paragraphs.Count && LooksLikeFlatTableParagraph(document.Paragraphs[i + 1]))
            {
                i++;
            }

            int end = i;
            var group = document.Paragraphs.Skip(start).Take(end - start + 1).ToList();
            var result = BuildFlatTable(group, recoveredTables.Count, start, end);
            if (result.Table == null)
                continue;

            recoveredTables.Add(result.Table);
            if (result.TrailingParagraphs.Count > 0)
            {
                document.Paragraphs.InsertRange(end + 1, result.TrailingParagraphs);
                i += result.TrailingParagraphs.Count;
            }
        }

        if (recoveredTables.Count == 0)
        {
            Log("RecoverFlatTables(method) no recovered tables");
            Log("RecoverFlatTables(method) DONE");
            return;
        }
        

        foreach (var table in recoveredTables)
        {
            AppendFollowingMetadataRows(document, table);
        }

        for (int index = 0; index < document.Paragraphs.Count; index++)
        {
            document.Paragraphs[index].Index = index;
        }

        foreach (var table in recoveredTables)
        {
            table.StartParagraphIndex = Math.Clamp(table.StartParagraphIndex, 0, document.Paragraphs.Count - 1);
            table.EndParagraphIndex = Math.Clamp(table.EndParagraphIndex, table.StartParagraphIndex, document.Paragraphs.Count - 1);
        }

        document.Tables = recoveredTables;

        foreach (var paragraph in document.Paragraphs)
        {
            ApplyRecoveredParagraphFormatting(paragraph, paragraph.Text);
        }
        Log("RecoverFlatTables(method) DONE");
    }

    private static void RecoverNestedTableSections(DocumentModel document)
    {
        if (document.Tables.Count == 0 || document.Paragraphs.Count == 0)
            return;

        var topLevelTables = document.Tables
            .Where(t => !t.IsNested)
            .OrderBy(t => t.StartParagraphIndex)
            .ToList();

        var rebuiltTables = new List<TableModel>();
        var consumedTables = new HashSet<TableModel>();

        for (int paragraphIndex = 0; paragraphIndex < document.Paragraphs.Count; paragraphIndex++)
        {
            var currentParagraph = document.Paragraphs[paragraphIndex];
            if (currentParagraph.Type == ParagraphType.TableCell || string.IsNullOrWhiteSpace(currentParagraph.Text))
                continue;

            int nextSectionTitleIndex = FindNextStandaloneParagraphIndex(document, paragraphIndex + 1);
            bool looksLikeNestedSectionTitle = LooksLikeNestedSectionTitle(currentParagraph.Text);
            var placeholderChild = TryBuildNestedPlaceholderTable(document, paragraphIndex, nextSectionTitleIndex);

            if (!looksLikeNestedSectionTitle && placeholderChild == null)
                continue;

            var nestedTable = topLevelTables.FirstOrDefault(table =>
                !consumedTables.Contains(table) &&
                table.StartParagraphIndex > paragraphIndex &&
                (nextSectionTitleIndex < 0 || table.StartParagraphIndex < nextSectionTitleIndex));

            if (nestedTable != null)
            {
                rebuiltTables.Add(BuildNestedSectionTable(nestedTable, currentParagraph.Index + 1));
                consumedTables.Add(nestedTable);
                continue;
            }

            if (placeholderChild != null)
            {
                rebuiltTables.Add(BuildNestedSectionTable(placeholderChild, currentParagraph.Index + 1));
            }
        }

        foreach (var table in topLevelTables)
        {
            if (!consumedTables.Contains(table))
            {
                rebuiltTables.Add(table);
            }
        }

        rebuiltTables = rebuiltTables
            .OrderBy(table => table.StartParagraphIndex)
            .ThenBy(table => table.EndParagraphIndex)
            .ToList();

        ReindexTopLevelTables(rebuiltTables);
        document.Tables = rebuiltTables;
    }

    private static int FindNextStandaloneParagraphIndex(DocumentModel document, int startIndex)
    {
        for (int index = startIndex; index < document.Paragraphs.Count; index++)
        {
            var paragraph = document.Paragraphs[index];
            if (paragraph.Type == ParagraphType.TableCell)
                continue;

            if (!string.IsNullOrWhiteSpace(paragraph.Text))
                return index;
        }

        return -1;
    }

    private static TableModel? TryBuildNestedPlaceholderTable(DocumentModel document, int titleIndex, int nextSectionTitleIndex)
    {
        int endIndex = nextSectionTitleIndex >= 0 ? nextSectionTitleIndex : document.Paragraphs.Count;
        var markerParagraphs = document.Paragraphs
            .Where(paragraph =>
                paragraph.Index > titleIndex &&
                paragraph.Index < endIndex &&
                paragraph.Type == ParagraphType.TableCell &&
                !string.IsNullOrEmpty(paragraph.RawText) &&
                paragraph.RawText.All(ch => ch == '\x07'))
            .ToList();

        if (markerParagraphs.Count == 0)
            return null;

        int markerCount = markerParagraphs.Sum(paragraph => paragraph.RawText.Count(ch => ch == '\x07'));
        if (markerCount < 4)
            return null;

        int columnCount = InferPlaceholderColumnCount(markerCount);
        int rowCount = Math.Max(1, markerCount / Math.Max(1, columnCount));
        var border = new BorderInfo { Style = BorderStyle.Single, Width = 4, Space = 0, Color = 0 };
        var childTable = new TableModel
        {
            StartParagraphIndex = markerParagraphs[0].Index,
            EndParagraphIndex = markerParagraphs[^1].Index,
            ColumnCount = columnCount,
            RowCount = rowCount,
            Properties = new TableProperties
            {
                PreferredWidth = 4680,
                BorderTop = border,
                BorderBottom = border,
                BorderLeft = border,
                BorderRight = border,
                BorderInsideH = border,
                BorderInsideV = border
            }
        };

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var row = new TableRowModel { Index = rowIndex };
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                row.Cells.Add(new TableCellModel
                {
                    Index = columnIndex,
                    RowIndex = rowIndex,
                    ColumnIndex = columnIndex,
                    Properties = new TableCellProperties { Width = 4680 / columnCount },
                    Paragraphs = new List<ParagraphModel> { new() }
                });
            }

            childTable.Rows.Add(row);
        }

        return childTable;
    }

    private static int InferPlaceholderColumnCount(int markerCount)
    {
        if (markerCount == 6)
            return 3;

        if (markerCount % 3 == 0 && markerCount / 3 >= 2)
            return 3;

        if (markerCount % 2 == 0)
            return 2;

        return Math.Min(3, markerCount);
    }

    private static TableModel BuildNestedSectionTable(TableModel childTable, int startParagraphIndex)
    {
        int columnCount = 2;
        int rowCount = 2;
        int cellWidth = 4680;
        var border = new BorderInfo { Style = BorderStyle.Single, Width = 4, Space = 0, Color = 0 };

        var parentTable = new TableModel
        {
            StartParagraphIndex = Math.Min(childTable.StartParagraphIndex, Math.Max(0, startParagraphIndex)),
            EndParagraphIndex = childTable.EndParagraphIndex,
            ColumnCount = columnCount,
            RowCount = rowCount,
            Properties = new TableProperties
            {
                PreferredWidth = 9360,
                BorderTop = border,
                BorderBottom = border,
                BorderLeft = border,
                BorderRight = border,
                BorderInsideH = border,
                BorderInsideV = border
            }
        };

        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var row = new TableRowModel { Index = rowIndex };
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var cell = new TableCellModel
                {
                    Index = columnIndex,
                    RowIndex = rowIndex,
                    ColumnIndex = columnIndex,
                    Properties = new TableCellProperties { Width = cellWidth }
                };

                if (rowIndex == 0 && columnIndex == 0)
                {
                    cell.Paragraphs.Add(new ParagraphModel
                    {
                        Type = ParagraphType.NestedTable,
                        NestedTable = childTable
                    });
                }
                else
                {
                    cell.Paragraphs.Add(new ParagraphModel());
                }

                row.Cells.Add(cell);
            }

            parentTable.Rows.Add(row);
        }

        return parentTable;
    }

    private static bool LooksLikeNestedSectionTitle(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return text.Contains("嵌套", StringComparison.Ordinal) ||
               text.Contains("nested", StringComparison.OrdinalIgnoreCase);
    }

    private static void ReindexTopLevelTables(List<TableModel> tables)
    {
        for (int index = 0; index < tables.Count; index++)
        {
            tables[index].Index = index;
            StampNestedParentIndex(tables[index], index);
        }
    }

    private static void StampNestedParentIndex(TableModel table, int parentIndex)
    {
        foreach (var row in table.Rows)
        {
            foreach (var cell in row.Cells)
            {
                foreach (var paragraph in cell.Paragraphs)
                {
                    if (paragraph.Type != ParagraphType.NestedTable || paragraph.NestedTable == null)
                        continue;

                    paragraph.NestedTable.ParentTableIndex = parentIndex;
                    StampNestedParentIndex(paragraph.NestedTable, paragraph.NestedTable.Index);
                }
            }
        }
    }

    private static void AppendFollowingMetadataRows(DocumentModel document, TableModel table)
    {
        if (table.ColumnCount < 2)
            return;

        int scanIndex = Math.Min(table.EndParagraphIndex + 1, document.Paragraphs.Count);
        var extraCells = new List<string>();
        var paragraphsToRemove = new List<ParagraphModel>();

        while (scanIndex < document.Paragraphs.Count)
        {
            var paragraph = document.Paragraphs[scanIndex];
            string text = paragraph.Text.Trim();

            if (string.IsNullOrWhiteSpace(text))
            {
                scanIndex++;
                continue;
            }

            if (LooksLikeSectionHeading(text))
            {
                scanIndex++;
                continue;
            }

            if (!LooksLikeMetadataCell(text))
                break;

            extraCells.Add(text);
            paragraphsToRemove.Add(paragraph);
            scanIndex++;
        }

        if (extraCells.Count < table.ColumnCount)
            return;

        for (int index = 0; index + table.ColumnCount <= extraCells.Count; index += table.ColumnCount)
        {
            var rowCells = extraCells.Skip(index).Take(table.ColumnCount).ToList();
            NormalizeMetadataPairOrder(rowCells);
            var recoveredCells = rowCells
                .Select((text, cellIndex) => new RecoveredCell
                {
                    Text = text,
                    SourceParagraph = paragraphsToRemove.ElementAtOrDefault(index + cellIndex)
                })
                .ToList();
            table.Rows.Add(BuildRecoveredRow(recoveredCells, table.Rows.Count, table.ColumnCount));
        }

        foreach (var paragraph in paragraphsToRemove)
        {
            document.Paragraphs.Remove(paragraph);
        }

        table.RowCount = table.Rows.Count;
    }

    private static bool LooksLikeFlatTableParagraph(ParagraphModel paragraph)
    {
        return paragraph.Type == ParagraphType.TableCell &&
               !string.IsNullOrEmpty(paragraph.RawText) &&
               paragraph.RawText.Contains('\x07');
    }

    private int EstimateColumnCount(IEnumerable<ParagraphModel> paragraphs)
    {
        int tapColumnCount = paragraphs
            .Select(GetTapColumnCount)
            .Where(count => count > 0)
            .DefaultIfEmpty(0)
            .Max();

        if (tapColumnCount >= 2)
            return tapColumnCount;

        var rowCandidates = paragraphs
            .SelectMany(GetRowCandidates)
            .ToList();

        if (rowCandidates.Count == 0)
            return 0;

        int maxCount = rowCandidates.Max(cells => cells.Count);
        if (maxCount <= 4)
            return maxCount;

        if (rowCandidates.Any(cells => cells.Count == 2))
            return 2;

        int labelLikeCells = rowCandidates
            .SelectMany(cells => cells)
            .Count(text => text.Contains('：') || text.Contains(':'));
        int totalCells = rowCandidates.Sum(cells => cells.Count);
        if (labelLikeCells >= Math.Max(4, totalCells - 1))
            return 2;

        return maxCount;
    }

    private (TableModel? Table, List<ParagraphModel> TrailingParagraphs) BuildFlatTable(
        List<ParagraphModel> group,
        int tableIndex,
        int startParagraphIndex,
        int endParagraphIndex)
    {
        int columnCount = EstimateColumnCount(group);
        if (columnCount < 2)
            return (null, new List<ParagraphModel>());

        var border = new BorderInfo { Style = BorderStyle.Single, Width = 4, Space = 0, Color = 0 };
        var table = new TableModel
        {
            Index = tableIndex,
            StartParagraphIndex = startParagraphIndex,
            EndParagraphIndex = endParagraphIndex,
            ColumnCount = columnCount,
            Properties = new TableProperties
            {
                BorderTop = border,
                BorderBottom = border,
                BorderLeft = border,
                BorderRight = border,
                BorderInsideH = border,
                BorderInsideV = border,
                PreferredWidth = 9360
            }
        };

        var trailingParagraphs = new List<ParagraphModel>();
        int rowIndex = 0;

        for (int paragraphIndex = 0; paragraphIndex < group.Count; paragraphIndex++)
        {
            var paragraph = group[paragraphIndex];
            var rowCandidates = BuildRecoveredCellRows(paragraph, GetRowCandidates(paragraph).ToList());
            if (rowCandidates.Count == 0)
                continue;

            bool isLastParagraph = paragraphIndex == group.Count - 1;
            for (int candidateIndex = 0; candidateIndex < rowCandidates.Count; candidateIndex++)
            {
                var cells = rowCandidates[candidateIndex];
                bool isLastCandidate = isLastParagraph && candidateIndex == rowCandidates.Count - 1;

                if (isLastCandidate)
                {
                    string trailingText = SplitTrailingText(ref cells, columnCount);
                    if (!string.IsNullOrWhiteSpace(trailingText))
                    {
                        trailingParagraphs.AddRange(BuildTrailingParagraphs(trailingText, paragraph));
                    }
                }

                while (cells.Count > columnCount)
                {
                    table.Rows.Add(BuildRecoveredRow(cells.Take(columnCount).ToList(), rowIndex++, columnCount));
                    cells = cells.Skip(columnCount).ToList();
                }

                if (cells.Count > 0)
                {
                    while (cells.Count < columnCount)
                    {
                        cells.Add(new RecoveredCell { Text = string.Empty, SourceParagraph = paragraph });
                    }

                    table.Rows.Add(BuildRecoveredRow(cells, rowIndex++, columnCount));
                }
            }
        }

        table.RowCount = table.Rows.Count;
        if (table.RowCount == 0)
            return (null, new List<ParagraphModel>());

        return (table, trailingParagraphs);
    }

    private static TableRowModel BuildRecoveredRow(List<RecoveredCell> cellTexts, int rowIndex, int columnCount)
    {
        var row = new TableRowModel { Index = rowIndex };
        int cellWidth = 9360 / Math.Max(1, columnCount);

        for (int columnIndex = 0; columnIndex < columnCount && columnIndex < cellTexts.Count; columnIndex++)
        {
            var recoveredCell = cellTexts[columnIndex];
            string cellText = recoveredCell.Text.Trim();
            var sourceParagraph = recoveredCell.SourceParagraph;
            var paragraph = new ParagraphModel
            {
                Type = ParagraphType.Normal,
                Properties = CloneParagraphProperties(sourceParagraph?.Properties),
                Runs = new List<RunModel>()
            };
            if (paragraph.Properties != null)
            {
                paragraph.Properties.KeepWithNext = false;
            }

            if (recoveredCell.SourceRuns.Count > 0)
            {
                paragraph.Runs.AddRange(recoveredCell.SourceRuns.Select(CloneRecoveredRun));
            }
            else if (!string.IsNullOrEmpty(cellText))
            {
                AddRecoveredRuns(paragraph, cellText, sourceParagraph?.Runs.FirstOrDefault()?.Properties);
            }

            ApplyRecoveredParagraphFormatting(paragraph, cellText);

            row.Cells.Add(new TableCellModel
            {
                Index = columnIndex,
                RowIndex = rowIndex,
                ColumnIndex = columnIndex,
                Paragraphs = new List<ParagraphModel> { paragraph },
                Properties = new TableCellProperties { Width = cellWidth }
            });
        }

        return row;
    }

    private bool TryConsumeCompactTableParagraph(ParagraphModel paragraph, TableContext ctx, TapBase? tapForParagraph)
    {
        if (!LooksLikeFlatTableParagraph(paragraph))
            return false;

        if (ctx.CurrentCellParagraphs.Count > 0 || ctx.CellsInCurrentRow.Count > 0)
            return false;

        var rowCandidates = BuildRecoveredCellRows(paragraph, GetRowCandidates(paragraph).ToList());

        if (rowCandidates.Count == 0)
            return false;

        bool encodesMultipleCells = rowCandidates.Count > 1 || rowCandidates.Any(cells => cells.Count > 1);
        if (!encodesMultipleCells)
            return false;

        foreach (var recoveredCells in rowCandidates)
        {
            ctx.CurrentRowTap = tapForParagraph;

            foreach (var recoveredCell in recoveredCells)
            {
                ctx.CurrentCellParagraphs.Add(BuildRecoveredParagraph(recoveredCell));
                FlushCurrentCell(ctx);
            }

            FlushCurrentRow(ctx);
        }

        return true;
    }

    private static ParagraphModel BuildRecoveredParagraph(RecoveredCell recoveredCell)
    {
        string cellText = recoveredCell.Text.Trim();
        var sourceParagraph = recoveredCell.SourceParagraph;
        var paragraph = new ParagraphModel
        {
            Type = ParagraphType.Normal,
            Properties = CloneParagraphProperties(sourceParagraph?.Properties),
            Runs = new List<RunModel>()
        };

        if (paragraph.Properties != null)
        {
            paragraph.Properties.KeepWithNext = false;
        }

        if (!string.IsNullOrEmpty(cellText))
        {
            if (recoveredCell.SourceRuns.Count > 0)
            {
                paragraph.Runs.AddRange(recoveredCell.SourceRuns.Select(CloneRecoveredRun));
            }
            else
            {
                AddRecoveredRuns(paragraph, cellText, sourceParagraph?.Runs.FirstOrDefault()?.Properties);
            }
        }

        ApplyRecoveredParagraphFormatting(paragraph, cellText);
        return paragraph;
    }

    private static IEnumerable<List<string>> GetRowCandidates(ParagraphModel paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.RawText))
            yield break;

        foreach (var rawLine in Regex.Split(paragraph.RawText, "(?:\r+|\x07{2,})"))
        {
            var cells = rawLine
                .Split('\x07')
                .Select(NormalizeFlatCellText)
                .ToList();

            while (cells.Count > 0 && string.IsNullOrWhiteSpace(cells[^1]))
                cells.RemoveAt(cells.Count - 1);

            if (cells.Count == 0)
                continue;

            cells = cells
                .Select(text => text)
                .ToList();

            if (cells.Count > 0)
                yield return cells;
        }
    }

    private int GetTapColumnCount(ParagraphModel paragraph)
    {
        var firstRun = paragraph.Runs.FirstOrDefault();
        if (firstRun == null)
            return 0;

        var pap = _fkpParser.GetPapAtCp(firstRun.CharacterPosition);
        var widths = pap?.Tap?.CellWidths;
        return widths?.Length ?? 0;
    }

    private static string NormalizeFlatCellText(string text)
    {
        return text
            .Replace("\r", "\n")
            .Trim('\r', '\n', '\t', ' ');
    }

    private static List<RecoveredCell> BuildRecoveredCells(ParagraphModel paragraph, List<string> cells)
    {
        var recoveredCells = cells
            .Select(text => new RecoveredCell { Text = text, SourceParagraph = paragraph })
            .ToList();

        if (paragraph.Runs.Count == 0)
            return recoveredCells;

        int runIndex = 0;
        int runTextOffset = 0;

        foreach (var cell in recoveredCells)
        {
            PopulateRecoveredCellRuns(cell, paragraph.Runs, ref runIndex, ref runTextOffset);
        }

        return recoveredCells;
    }

    private static List<List<RecoveredCell>> BuildRecoveredCellRows(ParagraphModel paragraph, List<List<string>> rowCandidates)
    {
        var recoveredRows = new List<List<RecoveredCell>>(rowCandidates.Count);
        int runIndex = 0;
        int runTextOffset = 0;

        foreach (var rowCandidate in rowCandidates)
        {
            var recoveredCells = rowCandidate
                .Select(text => new RecoveredCell { Text = text, SourceParagraph = paragraph })
                .ToList();

            foreach (var cell in recoveredCells)
            {
                PopulateRecoveredCellRuns(cell, paragraph.Runs, ref runIndex, ref runTextOffset);
            }

            recoveredRows.Add(recoveredCells);
        }

        return recoveredRows;
    }

    private static void PopulateRecoveredCellRuns(RecoveredCell recoveredCell, List<RunModel> sourceRuns, ref int runIndex, ref int runTextOffset)
    {
        var remainingText = recoveredCell.Text.Trim();
        if (string.IsNullOrEmpty(remainingText))
            return;

        while (runIndex < sourceRuns.Count && remainingText.Length > 0)
        {
            var sourceRun = sourceRuns[runIndex];
            var sourceText = sourceRun.Text ?? string.Empty;

            if (runTextOffset >= sourceText.Length)
            {
                runIndex++;
                runTextOffset = 0;
                continue;
            }

            var availableText = sourceText[runTextOffset..];
            if (availableText.Length == 0)
            {
                runIndex++;
                runTextOffset = 0;
                continue;
            }

            int commonPrefixLength = GetSharedPrefixLength(availableText, remainingText);
            if (commonPrefixLength == 0)
            {
                if (char.IsWhiteSpace(availableText[0]) || availableText[0] == '\x07')
                {
                    runTextOffset++;
                    continue;
                }

                break;
            }

            recoveredCell.SourceRuns.Add(CloneRecoveredRunSegment(sourceRun, runTextOffset, commonPrefixLength));
            runTextOffset += commonPrefixLength;
            remainingText = remainingText[commonPrefixLength..];

            if (runTextOffset >= sourceText.Length)
            {
                runIndex++;
                runTextOffset = 0;
            }
        }
    }

    private static int GetSharedPrefixLength(string availableText, string expectedText)
    {
        int maxLength = Math.Min(availableText.Length, expectedText.Length);
        int index = 0;
        while (index < maxLength && availableText[index] == expectedText[index])
        {
            index++;
        }

        return index;
    }

    private static bool LooksLikeMetadataCell(string text)
    {
        return text.Length <= 40 &&
               (text.Contains('：') || text.Contains(':')) &&
               !LooksLikeSectionHeading(text);
    }

    private static bool LooksLikeSectionHeading(string text)
    {
        return Regex.IsMatch(text, "^(第[一二三四五六七八九十百0-9]+条[：:].*|[0-9]+\\.[0-9].*)$");
    }

    private static void NormalizeMetadataPairOrder(List<string> rowCells)
    {
        if (rowCells.Count != 2)
            return;

        string firstLabel = GetMetadataLabel(rowCells[0]);
        string secondLabel = GetMetadataLabel(rowCells[1]);
        if (!string.Equals(firstLabel, secondLabel, StringComparison.Ordinal))
            return;

        if (rowCells[1].Length > rowCells[0].Length)
        {
            (rowCells[0], rowCells[1]) = (rowCells[1], rowCells[0]);
        }
    }

    private static string GetMetadataLabel(string text)
    {
        int separatorIndex = text.IndexOfAny(new[] { '：', ':' });
        return separatorIndex > 0 ? text[..separatorIndex] : text;
    }

    private static string SplitTrailingText(ref List<RecoveredCell> cells, int columnCount)
    {
        if (cells.Count == 0)
            return string.Empty;

        if (cells.Count > columnCount)
        {
            var trailingOverflow = cells.Skip(columnCount).ToList();
            cells = cells.Take(columnCount).ToList();
            return string.Join("\r", trailingOverflow.Select(cell => cell.Text));
        }

        if (cells.Count < columnCount)
            return string.Empty;

        var lastCell = cells[columnCount - 1].Text;
        var match = Regex.Match(lastCell, "(?<cell>.*?)(?<trail>(第[一二三四五六七八九十百0-9]+条[：:].*|[0-9]+\\.[0-9].*))$");
        if (!match.Success)
            return string.Empty;

        var cellText = match.Groups["cell"].Value.Trim();
        var trailingText = match.Groups["trail"].Value.Trim();
        if (string.IsNullOrWhiteSpace(cellText) || string.IsNullOrWhiteSpace(trailingText))
            return string.Empty;

        cells[columnCount - 1] = new RecoveredCell
        {
            Text = cellText,
            SourceParagraph = cells[columnCount - 1].SourceParagraph
        };
        return trailingText;
    }

    private static List<ParagraphModel> BuildTrailingParagraphs(string trailingText, ParagraphModel? sourceParagraph)
    {
        var paragraphs = new List<ParagraphModel>();
        if (string.IsNullOrWhiteSpace(trailingText))
            return paragraphs;

        foreach (var rawPart in trailingText.Split(new[] { '\r' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var part = rawPart.Trim();
            if (string.IsNullOrWhiteSpace(part))
                continue;

            var paragraph = new ParagraphModel
            {
                Type = ParagraphType.Normal,
                RawText = part,
                Properties = CloneParagraphProperties(sourceParagraph?.Properties),
                Runs = new List<RunModel>()
            };
            if (paragraph.Properties != null)
            {
                paragraph.Properties.KeepWithNext = false;
            }
            AddRecoveredRuns(paragraph, part, sourceParagraph?.Runs.FirstOrDefault()?.Properties);
            ApplyRecoveredParagraphFormatting(paragraph, part);
            paragraphs.Add(paragraph);
        }

        return paragraphs;
    }

    private static void AddRecoveredRuns(ParagraphModel paragraph, string text, RunProperties? sourceRunProps)
    {
        var baseRunProps = CloneRunProperties(sourceRunProps);
        int separatorIndex = text.IndexOfAny(new[] { '：', ':' });
        bool isMetadataCell = separatorIndex > 0 && separatorIndex < text.Length && !LooksLikeSectionHeading(text);

        if (isMetadataCell)
        {
            var labelProps = CloneRunProperties(baseRunProps) ?? new RunProperties();
            labelProps.IsBold = true;
            labelProps.IsBoldCs = true;

            paragraph.Runs.Add(new RunModel
            {
                Text = text[..(separatorIndex + 1)],
                Properties = labelProps
            });

            var valueText = text[(separatorIndex + 1)..];
            if (!string.IsNullOrEmpty(valueText))
            {
                RunProperties? valueProps = null;
                if (baseRunProps != null)
                {
                    valueProps = new RunProperties
                    {
                        FontIndex = baseRunProps.FontIndex,
                        FontName = baseRunProps.FontName,
                        FontSize = baseRunProps.FontSize,
                        FontSizeCs = baseRunProps.FontSizeCs,
                        Color = baseRunProps.Color,
                        BgColor = baseRunProps.BgColor,
                        Language = baseRunProps.Language,
                        LanguageAsia = baseRunProps.LanguageAsia,
                        LanguageCs = baseRunProps.LanguageCs,
                        HighlightColor = baseRunProps.HighlightColor,
                        RgbColor = baseRunProps.RgbColor,
                        HasRgbColor = baseRunProps.HasRgbColor,
                        CharacterSpacingAdjustment = baseRunProps.CharacterSpacingAdjustment,
                        CharacterScale = baseRunProps.CharacterScale,
                        SnapToGrid = baseRunProps.SnapToGrid
                    };
                }

                if (valueProps != null)
                {
                    valueProps.IsBold = false;
                    valueProps.IsBoldCs = false;
                }

                paragraph.Runs.Add(new RunModel
                {
                    Text = valueText,
                    Properties = valueProps
                });
            }

            return;
        }

        paragraph.Runs.Add(new RunModel
        {
            Text = text,
            Properties = baseRunProps
        });
    }

    private static RunModel CloneRecoveredRun(RunModel source)
    {
        return new RunModel
        {
            Text = source.Text,
            CharacterPosition = source.CharacterPosition,
            CharacterLength = source.CharacterLength,
            IsPicture = source.IsPicture,
            ImageIndex = source.ImageIndex,
            DisplayWidthTwips = source.DisplayWidthTwips,
            DisplayHeightTwips = source.DisplayHeightTwips,
            FcPic = source.FcPic,
            IsField = source.IsField,
            FieldCode = source.FieldCode,
            IsHyperlink = source.IsHyperlink,
            HyperlinkUrl = source.HyperlinkUrl,
            HyperlinkBookmark = source.HyperlinkBookmark,
            HyperlinkRelationshipId = source.HyperlinkRelationshipId,
            IsBookmark = source.IsBookmark,
            BookmarkName = source.BookmarkName,
            IsBookmarkStart = source.IsBookmarkStart,
            IsOle = source.IsOle,
            OleObjectId = source.OleObjectId,
            OleProgId = source.OleProgId,
            ImageRelationshipId = source.ImageRelationshipId,
            CropTop = source.CropTop,
            CropBottom = source.CropBottom,
            CropLeft = source.CropLeft,
            CropRight = source.CropRight,
            Properties = CloneRunProperties(source.Properties)
        };
    }

    private static RunModel CloneRecoveredRunSegment(RunModel source, int textOffset, int textLength)
    {
        var clone = CloneRecoveredRun(source);
        clone.Text = (source.Text ?? string.Empty).Substring(textOffset, textLength);
        clone.CharacterPosition = source.CharacterPosition + textOffset;
        clone.CharacterLength = textLength;
        return clone;
    }

    private static void ApplyRecoveredParagraphFormatting(ParagraphModel paragraph, string text)
    {
        if (!LooksLikeArticleHeading(text))
            return;

        paragraph.Properties ??= new ParagraphProperties();
        paragraph.Properties.KeepWithNext = true;
        paragraph.Properties.SpaceBefore = Math.Max(paragraph.Properties.SpaceBefore, 240);
        paragraph.Properties.SpaceAfter = Math.Max(paragraph.Properties.SpaceAfter, 240);

        foreach (var run in paragraph.Runs)
        {
            run.Properties ??= new RunProperties();
            run.Properties.IsBold = true;
            run.Properties.IsBoldCs = true;
        }
    }

    private static bool LooksLikeArticleHeading(string text)
    {
        return Regex.IsMatch(text, "^第[一二三四五六七八九十百0-9]+[条章节][：:、].*");
    }

    private static ParagraphProperties? CloneParagraphProperties(ParagraphProperties? source)
    {
        if (source == null)
            return null;

        return new ParagraphProperties
        {
            StyleIndex = source.StyleIndex,
            Alignment = source.Alignment,
            IndentLeft = source.IndentLeft,
            IndentLeftChars = source.IndentLeftChars,
            IndentRight = source.IndentRight,
            IndentRightChars = source.IndentRightChars,
            IndentFirstLine = source.IndentFirstLine,
            IndentFirstLineChars = source.IndentFirstLineChars,
            SpaceBefore = source.SpaceBefore,
            SpaceBeforeLines = source.SpaceBeforeLines,
            SpaceAfter = source.SpaceAfter,
            SpaceAfterLines = source.SpaceAfterLines,
            LineSpacing = source.LineSpacing,
            LineSpacingMultiple = source.LineSpacingMultiple,
            KeepWithNext = source.KeepWithNext,
            KeepTogether = source.KeepTogether,
            PageBreakBefore = source.PageBreakBefore,
            BorderTop = source.BorderTop,
            BorderBottom = source.BorderBottom,
            BorderLeft = source.BorderLeft,
            BorderRight = source.BorderRight,
            Shading = source.Shading,
            ListFormatId = source.ListFormatId,
            ListLevel = source.ListLevel,
            OutlineLevel = source.OutlineLevel,
            NumberFormat = source.NumberFormat,
            NumberText = source.NumberText,
            SnapToGrid = source.SnapToGrid,
            AutoSpaceDe = source.AutoSpaceDe,
            AutoSpaceDn = source.AutoSpaceDn,
            WordWrap = source.WordWrap,
            Kinsoku = source.Kinsoku,
            OverflowPunct = source.OverflowPunct,
            TopLinePunct = source.TopLinePunct
        };
    }

    private static RunProperties? CloneRunProperties(RunProperties? source)
    {
        if (source == null)
            return null;

        return new RunProperties
        {
            FontIndex = source.FontIndex,
            FontName = source.FontName,
            FontSize = source.FontSize,
            FontSizeCs = source.FontSizeCs,
            IsBold = source.IsBold,
            IsBoldCs = source.IsBoldCs,
            IsItalic = source.IsItalic,
            IsItalicCs = source.IsItalicCs,
            IsUnderline = source.IsUnderline,
            UnderlineType = source.UnderlineType,
            IsStrikeThrough = source.IsStrikeThrough,
            IsDoubleStrikeThrough = source.IsDoubleStrikeThrough,
            IsSmallCaps = source.IsSmallCaps,
            IsAllCaps = source.IsAllCaps,
            IsHidden = source.IsHidden,
            IsSuperscript = source.IsSuperscript,
            IsSubscript = source.IsSubscript,
            Color = source.Color,
            BgColor = source.BgColor,
            CharacterSpacingAdjustment = source.CharacterSpacingAdjustment,
            Language = source.Language,
            LanguageAsia = source.LanguageAsia,
            LanguageCs = source.LanguageCs,
            HighlightColor = source.HighlightColor,
            RgbColor = source.RgbColor,
            HasRgbColor = source.HasRgbColor,
            IsOutline = source.IsOutline,
            IsShadow = source.IsShadow,
            IsEmboss = source.IsEmboss,
            IsImprint = source.IsImprint,
            Kerning = source.Kerning,
            Position = source.Position,
            CharacterScale = source.CharacterScale,
            SnapToGrid = source.SnapToGrid,
            RubyText = source.RubyText,
            IsDeleted = source.IsDeleted,
            IsInserted = source.IsInserted,
            AuthorIndexDel = source.AuthorIndexDel,
            AuthorIndexIns = source.AuthorIndexIns,
            DateDel = source.DateDel,
            DateIns = source.DateIns
        };
    }

    private void FlushCurrentCell(TableContext ctx)
    {
        Log($"FlushCurrentCell START level={ctx.Level} cellsInRow={ctx.CellsInCurrentRow.Count} currentCellParas={ctx.CurrentCellParagraphs.Count}");
        if (ctx.CurrentCellParagraphs.Count == 0 && ctx.CellsInCurrentRow.Count == 0)
        {
            Log($"FlushCurrentCell SKIP level={ctx.Level} reason=empty-context");
            return;
        }

        if (ctx.CurrentCellParagraphs.Count == 0 && ctx.CellsInCurrentRow.Count > 0)
        {
            // Empty cell? placeholder for future logic
        }

        var cellModel = new TableCellModel
        {
            Index = ctx.CellsInCurrentRow.Count,
            RowIndex = ctx.RowIndex,
            ColumnIndex = ctx.CellsInCurrentRow.Count,
            Paragraphs = new List<ParagraphModel>(ctx.CurrentCellParagraphs)
        };

        // Apply cell properties from TAP
        if (ctx.CurrentRowTap != null)
        {
            var colIdx = cellModel.ColumnIndex;
            
            // Cell width
            if (ctx.CurrentRowTap.CellWidths != null && ctx.CurrentRowTap.CellWidths.Length > colIdx)
            {
                cellModel.Properties ??= new TableCellProperties();
                cellModel.Properties.Width = ctx.CurrentRowTap.CellWidths[colIdx];
            }
            
            // Cell borders from TAP's cell borders if available
            if (ctx.CurrentRowTap.CellBorders != null && ctx.CurrentRowTap.CellBorders.Length > colIdx)
            {
                var cellBorders = ctx.CurrentRowTap.CellBorders[colIdx];
                if (cellBorders != null)
                {
                    cellModel.Properties ??= new TableCellProperties();
                    cellModel.Properties.BorderTop = cellBorders.Top;
                    cellModel.Properties.BorderBottom = cellBorders.Bottom;
                    cellModel.Properties.BorderLeft = cellBorders.Left;
                    cellModel.Properties.BorderRight = cellBorders.Right;
                }
            }
            
            // Cell shading from TAP if available
            if (ctx.CurrentRowTap.CellShadings != null && ctx.CurrentRowTap.CellShadings.Length > colIdx)
            {
                var shading = ctx.CurrentRowTap.CellShadings[colIdx];
                if (shading != null)
                {
                    cellModel.Properties ??= new TableCellProperties();
                    cellModel.Properties.Shading = shading;
                }
            }
            
            // Cell vertical alignment from TAP if available
            if (ctx.CurrentRowTap.CellVerticalAlignments != null && ctx.CurrentRowTap.CellVerticalAlignments.Length > colIdx)
            {
                cellModel.Properties ??= new TableCellProperties();
                cellModel.Properties.VerticalAlignment = (VerticalAlignment)ctx.CurrentRowTap.CellVerticalAlignments[colIdx];
            }
        }

        ctx.CellsInCurrentRow.Add(cellModel);
        ctx.CurrentCellParagraphs.Clear();
        Log($"FlushCurrentCell DONE level={ctx.Level} nowCellsInRow={ctx.CellsInCurrentRow.Count}");
    }

    private void FlushCurrentRow(TableContext ctx)
    {
        Log($"FlushCurrentRow START level={ctx.Level} cellsInRow={ctx.CellsInCurrentRow.Count}");
        if (ctx.CellsInCurrentRow.Count == 0) return;

        var row = new TableRowModel
        {
            Index = ctx.RowIndex++,
            Cells = new List<TableCellModel>(ctx.CellsInCurrentRow)
        };

        if (ctx.CurrentRowTap != null)
        {
            row.Properties ??= new TableRowProperties();
            if (ctx.CurrentRowTap.RowHeight > 0)
            {
                row.Properties.Height = ctx.CurrentRowTap.RowHeight;
                row.Properties.HeightIsExact = ctx.CurrentRowTap.HeightIsExact;
            }
            if (ctx.CurrentRowTap.IsHeaderRow)
            {
                row.Properties.IsHeaderRow = true;
            }
            row.Properties.AllowBreakAcrossPages = !ctx.CurrentRowTap.CantSplit;
        }

        ctx.Table.Rows.Add(row);
        ctx.RowTaps.Add(ctx.CurrentRowTap);
        ctx.CellsInCurrentRow.Clear();
        ctx.CurrentRowTap = null;
        Log($"FlushCurrentRow DONE level={ctx.Level} tableRows={ctx.Table.Rows.Count}");
    }

    private void FinalizeTable(TableContext ctx)
    {
        Log($"FinalizeTable START level={ctx.Level} rows={ctx.Table.Rows.Count}");
        var table = ctx.Table;
        if (table.Rows.Count == 0) return;

        var effectiveRowTaps = new List<TapBase?>(ctx.RowTaps);
        int? rebuiltEndParagraphIndex = null;
        int tapColumnCount = effectiveRowTaps
            .Select(tap => tap?.CellWidths?.Length ?? 0)
            .DefaultIfEmpty(0)
            .Max();
        var widthTemplate = GetWidthTemplate(effectiveRowTaps, tapColumnCount);

        if (tapColumnCount < 2)
        {
            tapColumnCount = InferSequentialSingleCellColumnCount(table);
        }

        if (TryRebuildCalendarMonthTable(table, effectiveRowTaps, tapColumnCount, out var rebuiltCalendarRowTaps, out var calendarEndParagraphIndex))
        {
            effectiveRowTaps = rebuiltCalendarRowTaps;
            rebuiltEndParagraphIndex = calendarEndParagraphIndex;
            widthTemplate ??= GetWidthTemplate(effectiveRowTaps, tapColumnCount);
        }
        else if (TryRebuildSequentialSingleCellTable(table, effectiveRowTaps, tapColumnCount, out var rebuiltRowTaps))
        {
            effectiveRowTaps = rebuiltRowTaps;
            widthTemplate ??= GetWidthTemplate(effectiveRowTaps, tapColumnCount);
        }

        table.EndParagraphIndex = ctx.LastTableParagraphIndex;
        if (rebuiltEndParagraphIndex.HasValue)
        {
            table.EndParagraphIndex = Math.Min(table.EndParagraphIndex, rebuiltEndParagraphIndex.Value);
        }
        if (table.EndParagraphIndex < table.StartParagraphIndex)
        {
            table.EndParagraphIndex = Math.Max(table.StartParagraphIndex + 1, 1);
            Log($"FinalizeTable: Corrected EndParagraphIndex to {table.EndParagraphIndex}");
        }
        
        table.RowCount = table.Rows.Count;
        if (table.RowCount > MaxRowsPerTable)
        {
            table.Rows = table.Rows.Take(MaxRowsPerTable).ToList();
            table.RowCount = table.Rows.Count;
        }
        table.ColumnCount = Math.Max(
            table.Rows.Max(r => r.Cells.Select(cell => cell.ColumnIndex + Math.Max(1, cell.ColumnSpan)).DefaultIfEmpty(r.Cells.Count).Max()),
            tapColumnCount);

        // Fix cell indices - ensure RowIndex and ColumnIndex are correct
        for (int rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            var row = table.Rows[rowIdx];
            row.Index = rowIdx;
            for (int colIdx = 0; colIdx < row.Cells.Count; colIdx++)
            {
                var cell = row.Cells[colIdx];
                cell.RowIndex = rowIdx;
                cell.ColumnIndex = colIdx;
                cell.Index = colIdx;
            }
        }

        bool forceWidthTemplate = false;
        if (widthTemplate == null && TryInferCalendarMonthWidthTemplate(table, out var inferredWidthTemplate))
        {
            widthTemplate = inferredWidthTemplate;
            forceWidthTemplate = true;
        }

        ApplyWidthTemplate(table, widthTemplate, forceWidthTemplate);

        // Only set header row when the TAP data explicitly flags it.
        // Do NOT force all first rows to be headers — that's wrong for most tables.
        var firstRow = table.Rows.FirstOrDefault();
        var firstTap = effectiveRowTaps.Count > 0 ? effectiveRowTaps[0] : null;
        if (firstRow != null && firstTap != null && firstTap.IsHeaderRow)
        {
            firstRow.Properties ??= new TableRowProperties();
            firstRow.Properties.IsHeaderRow = true;
        }

        // Apply Spans
        bool hasTapMergeInfo = effectiveRowTaps.Any(t => t?.CellMerges != null);
        if (hasTapMergeInfo && table.ColumnCount > 0)
        {
            for (int col = 0; col < table.ColumnCount; col++)
            {
                int row = 0;
                while (row < table.Rows.Count)
                {
                    var startCell = GetCell(table, row, col);
                    if (startCell == null) { row++; continue; }

                    var tap = row < effectiveRowTaps.Count ? effectiveRowTaps[row] : null;
                    var flags = tap?.CellMerges != null && col < tap.CellMerges.Length ? tap.CellMerges[col] : null;
                    if (flags == null || !flags.VertFirst) { row++; continue; }

                    int span = 1;
                    int nextRow = row + 1;
                    while (nextRow < table.Rows.Count)
                    {
                        var nextTap = nextRow < effectiveRowTaps.Count ? effectiveRowTaps[nextRow] : null;
                        var nextFlags = nextTap?.CellMerges != null && col < nextTap.CellMerges.Length ? nextTap.CellMerges[col] : null;
                        if (nextFlags == null || !nextFlags.VertMerged) break;
                        span++;
                        nextRow++;
                    }

                    if (span > 1) { startCell.RowSpan = span; row += span; }
                    else { row++; }
                }
            }

            for (int row = 0; row < table.Rows.Count; row++)
            {
                var tap = row < effectiveRowTaps.Count ? effectiveRowTaps[row] : null;
                var mergeArray = tap?.CellMerges;
                if (mergeArray == null || mergeArray.Length == 0) continue;

                int col = 0;
                while (col < table.ColumnCount)
                {
                    var cell = GetCell(table, row, col);
                    if (cell == null) { col++; continue; }

                    var flags = col < mergeArray.Length ? mergeArray[col] : null;
                    if (flags == null || !flags.HorizFirst) { col++; continue; }

                    int span = 1;
                    int nextCol = col + 1;
                    while (nextCol < table.ColumnCount)
                    {
                        var nextFlags = nextCol < mergeArray.Length ? mergeArray[nextCol] : null;
                        if (nextFlags == null || !nextFlags.HorizMerged) break;
                        span++;
                        nextCol++;
                    }

                    if (span > 1) { cell.ColumnSpan = span; col += span; }
                    else { col++; }
                }
            }
        }
        Log($"FinalizeTable DONE level={ctx.Level} rows={table.RowCount} cols={table.ColumnCount}");
    }

    private static bool TryRebuildCalendarMonthTable(
        TableModel table,
        List<TapBase?> rowTaps,
        int tapColumnCount,
        out List<TapBase?> rebuiltRowTaps,
        out int? rebuiltEndParagraphIndex)
    {
        rebuiltRowTaps = rowTaps;
        rebuiltEndParagraphIndex = null;

        if (tapColumnCount != 13 || table.Rows.Count < 39 || table.Rows.Any(row => row.Cells.Count != 1))
            return false;

        var flattenedCells = table.Rows.Select(row => row.Cells[0]).ToList();
        var texts = flattenedCells
            .Select(cell => cell.Paragraphs.FirstOrDefault()?.Text?.Trim() ?? string.Empty)
            .ToList();

        if (!TryParseMonthYear(texts.FirstOrDefault(), out int month, out int year))
            return false;

        string[] expectedDays = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        if (texts.Count < 8 || !texts.Skip(1).Take(expectedDays.Length).SequenceEqual(expectedDays, StringComparer.OrdinalIgnoreCase))
            return false;

        int daysInMonth = DateTime.DaysInMonth(year, month);
        int requiredCells = 1 + expectedDays.Length + daysInMonth;
        if (flattenedCells.Count < requiredCells)
            return false;

        var dateTexts = texts.Skip(1 + expectedDays.Length).Take(daysInMonth).ToList();
        if (!dateTexts.Select((text, index) => string.Equals(text, (index + 1).ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal)).All(match => match))
            return false;

        var rebuiltRows = new List<TableRowModel>();
        var effectiveTaps = new List<TapBase?>();

        var titleCell = flattenedCells[0];
        ConfigureRebuiltCell(titleCell, rowTaps.ElementAtOrDefault(0), 0, tapColumnCount);
        titleCell.ColumnSpan = tapColumnCount;
        rebuiltRows.Add(new TableRowModel
        {
            Index = 0,
            Properties = table.Rows[0].Properties,
            Cells = new List<TableCellModel> { titleCell }
        });
        effectiveTaps.Add(rowTaps.ElementAtOrDefault(0));

        var headerRow = CreateCalendarRow(1, rowTaps.ElementAtOrDefault(1));
        for (int dayIndex = 0; dayIndex < expectedDays.Length; dayIndex++)
        {
            int cellIndex = 1 + dayIndex;
            int columnIndex = dayIndex * 2;
            var headerCell = flattenedCells[cellIndex];
            ConfigureRebuiltCell(headerCell, rowTaps.ElementAtOrDefault(cellIndex), columnIndex, 1);
            headerRow.Cells[columnIndex] = headerCell;
        }
        rebuiltRows.Add(headerRow);
        effectiveTaps.Add(rowTaps.ElementAtOrDefault(1));

        int firstDayOfWeek = (int)new DateTime(year, month, 1).DayOfWeek;
        int dateStartIndex = 1 + expectedDays.Length;
        for (int weekIndex = 0; weekIndex < 6; weekIndex++)
        {
            int dateRowIndex = 2 + (weekIndex * 2);
            var dateRowTap = rowTaps.ElementAtOrDefault(Math.Min(dateStartIndex + (weekIndex * 7), rowTaps.Count - 1));
            var dateRow = CreateCalendarRow(dateRowIndex, dateRowTap);

            rebuiltRows.Add(dateRow);
            effectiveTaps.Add(dateRowTap);

            if (dateRowIndex < 12)
            {
                rebuiltRows.Add(CreateCalendarRow(dateRowIndex + 1, dateRowTap));
                effectiveTaps.Add(dateRowTap);
            }
        }

        for (int day = 1; day <= daysInMonth; day++)
        {
            int weekIndex = (firstDayOfWeek + day - 1) / 7;
            int dayOfWeekIndex = (firstDayOfWeek + day - 1) % 7;
            int rowIndex = 2 + (weekIndex * 2);
            int columnIndex = dayOfWeekIndex * 2;
            int sourceCellIndex = dateStartIndex + (day - 1);

            var dateCell = flattenedCells[sourceCellIndex];
            ConfigureRebuiltCell(dateCell, rowTaps.ElementAtOrDefault(sourceCellIndex), columnIndex, 1);
            rebuiltRows[rowIndex].Cells[columnIndex] = dateCell;
        }

        table.Rows = rebuiltRows;
        rebuiltRowTaps = effectiveTaps;
        rebuiltEndParagraphIndex = flattenedCells
            .Take(requiredCells)
            .SelectMany(cell => cell.Paragraphs)
            .Select(paragraph => paragraph.Index)
            .DefaultIfEmpty(table.EndParagraphIndex)
            .Max();
        return true;
    }

    private static bool TryRebuildSequentialSingleCellTable(
        TableModel table,
        List<TapBase?> rowTaps,
        int tapColumnCount,
        out List<TapBase?> rebuiltRowTaps)
    {
        rebuiltRowTaps = rowTaps;

        if (tapColumnCount < 2 || table.Rows.Count <= tapColumnCount)
            return false;

        if (table.Rows.Any(row => row.Cells.Count != 1))
            return false;

        var flattenedCells = table.Rows.Select(row => row.Cells[0]).ToList();
        if (flattenedCells.Count <= tapColumnCount)
            return false;

        int leadingSpan = GetLeadingHorizontalSpan(rowTaps.FirstOrDefault(), tapColumnCount);
        bool hasMergedLeadingCell = leadingSpan == tapColumnCount && flattenedCells.Count > tapColumnCount;
        if (!hasMergedLeadingCell && LooksLikeMergedCalendarTitle(flattenedCells, tapColumnCount))
        {
            hasMergedLeadingCell = true;
        }
        var rebuiltRows = new List<TableRowModel>();
        var effectiveTaps = new List<TapBase?>();
        int cellCursor = 0;

        if (hasMergedLeadingCell)
        {
            var titleTap = rowTaps[0];
            var titleCell = flattenedCells[0];
            ConfigureRebuiltCell(titleCell, titleTap, 0, tapColumnCount);
            titleCell.ColumnSpan = tapColumnCount;

            rebuiltRows.Add(new TableRowModel
            {
                Index = 0,
                Properties = table.Rows[0].Properties,
                Cells = new List<TableCellModel> { titleCell }
            });
            effectiveTaps.Add(titleTap);
            cellCursor = 1;
        }

        while (cellCursor < flattenedCells.Count)
        {
            int logicalRowIndex = rebuiltRows.Count;
            int sourceRowIndex = Math.Min(cellCursor, table.Rows.Count - 1);
            var rowTap = sourceRowIndex < rowTaps.Count ? rowTaps[sourceRowIndex] : rowTaps.LastOrDefault();
            var row = new TableRowModel
            {
                Index = logicalRowIndex,
                Properties = table.Rows[sourceRowIndex].Properties,
                Cells = new List<TableCellModel>()
            };

            for (int columnIndex = 0; columnIndex < tapColumnCount && cellCursor < flattenedCells.Count; columnIndex++, cellCursor++)
            {
                var cell = flattenedCells[cellCursor];
                ConfigureRebuiltCell(cell, rowTap, columnIndex, 1);
                row.Cells.Add(cell);
            }

            while (row.Cells.Count < tapColumnCount)
            {
                int columnIndex = row.Cells.Count;
                var emptyCell = new TableCellModel
                {
                    Paragraphs = new List<ParagraphModel> { new() },
                    Properties = new TableCellProperties()
                };
                ConfigureRebuiltCell(emptyCell, rowTap, columnIndex, 1);
                row.Cells.Add(emptyCell);
            }

            rebuiltRows.Add(row);
            effectiveTaps.Add(rowTap);
        }

        table.Rows = rebuiltRows;
        rebuiltRowTaps = effectiveTaps;
        return true;
    }

    private static bool TryInferCalendarMonthWidthTemplate(TableModel table, out int[]? widthTemplate)
    {
        widthTemplate = null;

        if (!TryGetCalendarMonthColumnPattern(table, out var contentColumns, out var separatorColumns))
            return false;

        int preferredWidth = table.Properties?.PreferredWidth > 0 ? table.Properties.PreferredWidth : 0;
        if (preferredWidth <= 0)
        {
            preferredWidth = table.Rows[0].Cells.FirstOrDefault()?.Properties?.Width ?? 0;
        }

        if (preferredWidth <= 0 || contentColumns.Count == 0)
            return false;

        int separatorWidth = InferCalendarSeparatorWidth(table, preferredWidth, contentColumns.Count, separatorColumns.Count);
        int contentWidth = Math.Max(1, (preferredWidth - (separatorColumns.Count * separatorWidth)) / contentColumns.Count);
        int remainingWidth = preferredWidth - (contentWidth * contentColumns.Count) - (separatorWidth * separatorColumns.Count);

        var inferred = new int[table.ColumnCount];
        foreach (var column in contentColumns)
        {
            inferred[column] = contentWidth;
        }

        foreach (var column in separatorColumns)
        {
            inferred[column] = separatorWidth;
        }

        var distributionColumns = separatorColumns.Count > 0 ? separatorColumns : contentColumns;
        for (int index = 0; index < distributionColumns.Count && remainingWidth > 0; index++, remainingWidth--)
        {
            inferred[distributionColumns[index]]++;
        }

        widthTemplate = inferred;
        return true;
    }

    private static bool TryGetCalendarMonthColumnPattern(TableModel table, out List<int> contentColumns, out List<int> separatorColumns)
    {
        contentColumns = new List<int>();
        separatorColumns = new List<int>();

        if (table.ColumnCount != 13 || table.Rows.Count != 13)
            return false;

        var title = table.Rows[0].Cells.FirstOrDefault()?.Paragraphs.FirstOrDefault()?.Text;
        if (!LooksLikeMonthYear(title))
            return false;

        if (table.Rows.Count < 2 || table.Rows[1].Cells.Count != 13)
            return false;

        string[] expectedDays = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        var headerTexts = table.Rows[1].Cells.Select(cell => cell.Paragraphs.FirstOrDefault()?.Text ?? string.Empty).ToList();
        if (!headerTexts.Where((_, index) => index % 2 == 0).SequenceEqual(expectedDays, StringComparer.OrdinalIgnoreCase))
            return false;

        if (headerTexts.Where((_, index) => index % 2 == 1).Any(text => !string.IsNullOrWhiteSpace(text)))
            return false;

        for (int index = 0; index < table.ColumnCount; index++)
        {
            if (index % 2 == 0)
            {
                contentColumns.Add(index);
            }
            else
            {
                separatorColumns.Add(index);
            }
        }

        return true;
    }

    private static int InferCalendarSeparatorWidth(TableModel table, int preferredWidth, int contentColumnCount, int separatorColumnCount)
    {
        if (separatorColumnCount == 0)
            return 0;

        if ((table.Properties?.CellSpacing ?? 0) > 0)
        {
            return Math.Clamp(table.Properties!.CellSpacing, 1, preferredWidth / Math.Max(1, table.ColumnCount));
        }

        int defaultSeparatorWidth = Math.Max(1, preferredWidth / Math.Max(1, (contentColumnCount * 4) + separatorColumnCount));
        int maxSeparatorWidth = Math.Max(1, (preferredWidth - contentColumnCount) / Math.Max(1, contentColumnCount + separatorColumnCount));
        return Math.Min(defaultSeparatorWidth, maxSeparatorWidth);
    }

    private static int InferSequentialSingleCellColumnCount(TableModel table)
    {
        if (table.Rows.Count < 10 || table.Rows.Any(row => row.Cells.Count != 1))
            return 0;

        var texts = table.Rows
            .Select(row => row.Cells[0].Paragraphs.FirstOrDefault()?.Text?.Trim() ?? string.Empty)
            .ToList();

        if (texts.Count < 8 || !LooksLikeMonthYear(texts[0]))
            return 0;

        string[] expectedDays = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        if (!texts.Skip(1).Take(expectedDays.Length).SequenceEqual(expectedDays, StringComparer.OrdinalIgnoreCase))
            return 0;

        return 13;
    }

    private static bool LooksLikeMergedCalendarTitle(List<TableCellModel> flattenedCells, int tapColumnCount)
    {
        if (tapColumnCount != 13 || flattenedCells.Count <= tapColumnCount)
            return false;

        var title = flattenedCells[0].Paragraphs.FirstOrDefault()?.Text?.Trim();
        return LooksLikeMonthYear(title);
    }

    private static bool LooksLikeMonthYear(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        return Regex.IsMatch(text, "^(January|February|March|April|May|June|July|August|September|October|November|December)\\s+\\d{4}$", RegexOptions.IgnoreCase);
    }

    private static bool TryParseMonthYear(string? text, out int month, out int year)
    {
        month = 0;
        year = 0;

        if (string.IsNullOrWhiteSpace(text))
            return false;

        if (!DateTime.TryParseExact(text.Trim(), "MMMM yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
            return false;

        month = parsed.Month;
        year = parsed.Year;
        return true;
    }

    private static TableRowModel CreateCalendarRow(int rowIndex, TapBase? rowTap)
    {
        var row = new TableRowModel
        {
            Index = rowIndex,
            Properties = rowTap == null ? null : new TableRowProperties
            {
                Height = rowTap.RowHeight,
                HeightIsExact = rowTap.HeightIsExact,
                IsHeaderRow = rowTap.IsHeaderRow,
                AllowBreakAcrossPages = !rowTap.CantSplit
            }
        };

        for (int columnIndex = 0; columnIndex < 13; columnIndex++)
        {
            var emptyCell = new TableCellModel
            {
                Paragraphs = new List<ParagraphModel> { new() },
                Properties = new TableCellProperties()
            };
            ConfigureRebuiltCell(emptyCell, rowTap, columnIndex, 1);
            row.Cells.Add(emptyCell);
        }

        return row;
    }

    private static void ConfigureRebuiltCell(TableCellModel cell, TapBase? rowTap, int columnIndex, int columnSpan)
    {
        cell.Index = columnIndex;
        cell.ColumnIndex = columnIndex;
        cell.ColumnSpan = Math.Max(1, columnSpan);
        cell.RowSpan = Math.Max(1, cell.RowSpan);

        if (rowTap?.CellWidths == null || rowTap.CellWidths.Length <= columnIndex)
            return;

        cell.Properties ??= new TableCellProperties();

        int width = 0;
        for (int offset = 0; offset < cell.ColumnSpan && columnIndex + offset < rowTap.CellWidths.Length; offset++)
        {
            width += rowTap.CellWidths[columnIndex + offset];
        }

        if (width > 0)
        {
            cell.Properties.Width = width;
        }
    }

    private static int[]? GetWidthTemplate(List<TapBase?> rowTaps, int tapColumnCount)
    {
        if (tapColumnCount <= 0)
            return null;

        return rowTaps
            .Select(tap => tap?.CellWidths)
            .Where(widths => widths != null && widths.Length >= tapColumnCount)
            .Select(widths => widths!.Take(tapColumnCount).ToArray())
            .Where(widths => widths.Any(width => width > 0))
            .OrderByDescending(widths => widths.Count(width => width > 0))
            .ThenByDescending(widths => widths.Sum())
            .FirstOrDefault();
    }

    private static void ApplyWidthTemplate(TableModel table, int[]? widthTemplate, bool overwriteExistingWidths = false)
    {
        if (widthTemplate == null || widthTemplate.Length == 0)
            return;

        foreach (var row in table.Rows)
        {
            foreach (var cell in row.Cells)
            {
                if (!overwriteExistingWidths && cell.Properties?.Width > 0)
                    continue;

                int startColumn = Math.Clamp(cell.ColumnIndex, 0, widthTemplate.Length - 1);
                int endColumn = Math.Min(widthTemplate.Length, startColumn + Math.Max(1, cell.ColumnSpan));
                int width = 0;

                for (int column = startColumn; column < endColumn; column++)
                {
                    width += widthTemplate[column];
                }

                if (width <= 0)
                    continue;

                cell.Properties ??= new TableCellProperties();
                cell.Properties.Width = width;
            }
        }

        int preferredWidth = widthTemplate.Sum();
        if (preferredWidth <= 0)
            return;

        table.Properties ??= new TableProperties();
        if (table.Properties.PreferredWidth <= 0)
        {
            table.Properties.PreferredWidth = preferredWidth;
        }
    }

    private static int GetLeadingHorizontalSpan(TapBase? rowTap, int tapColumnCount)
    {
        var merges = rowTap?.CellMerges;
        if (merges == null || merges.Length == 0 || !merges[0].HorizFirst)
            return 0;

        int span = 1;
        for (int index = 1; index < merges.Length && index < tapColumnCount; index++)
        {
            if (!merges[index].HorizMerged)
                break;

            span++;
        }

        return span;
    }

    private static TableCellModel? GetCell(TableModel table, int rowIndex, int columnIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count) return null;
        var row = table.Rows[rowIndex];
        if (columnIndex < 0 || columnIndex >= row.Cells.Count) return null;
        return row.Cells[columnIndex];
    }

    private static bool CellHasContent(TableCellModel cell)
    {
        foreach (var para in cell.Paragraphs)
        {
            if (para.Runs.Any(r => !string.IsNullOrEmpty(r.Text) || r.IsPicture)) return true;
        }
        return false;
    }

    

    private static void Log(string message)
    {
        try
        {
            var line = $"{DateTime.UtcNow:O} {message}{Environment.NewLine}";
            File.AppendAllText(_debugLogPath, line);
        }
        catch
        {
            // Swallow logging errors to avoid affecting parsing.
        }
    }

    // Mark end of ParseTables explicitly to help trace long runs
    // (a separate statement rather than inline so it's easy to find in the log).
    // Note: this will be written by ParseTables before returning.

    private class TableContext
    {
        public int Level { get; set; }
        public TableModel Table { get; set; } = new();
        public int RowIndex { get; set; }
        public List<TableCellModel> CellsInCurrentRow { get; } = new();
        public List<ParagraphModel> CurrentCellParagraphs { get; } = new();
        public TapBase? CurrentRowTap { get; set; }
        public List<TapBase?> RowTaps { get; } = new();
        public int LastTableParagraphIndex { get; set; }
    }
}
