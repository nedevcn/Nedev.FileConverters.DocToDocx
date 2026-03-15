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
    public List<ParagraphModel> SourceParagraphs { get; } = new();
}

/// <summary>
/// Extracts tables (including nested tables) by interpreting paragraph nesting levels
/// and cell boundaries (\x07).
/// </summary>
public class TableReader
{
    private const int MaxNestingDepth = 20;
    private const int MaxRowsPerTable = 10000;
    private static readonly string _debugLogBaseName = "table_reader_debug";

    // Compiled regex patterns for better performance
    private static readonly Regex CellBoundaryRegex = new("\x07+", RegexOptions.Compiled);
    private static readonly Regex MultiCellBoundaryRegex = new("\x07{2,}", RegexOptions.Compiled);
    private static readonly Regex TrailingBoundaryRegex = new("\x07+$", RegexOptions.Compiled);
    private static readonly Regex LineBreakRegex = new("(?:\r+|\x07{2,})", RegexOptions.Compiled);
    private static readonly Regex SectionTitleRegex = new("(?<cell>.*?)(?<trail>(第[一二三四五六七八九十百0-9]+条[：:].*|[0-9]+\\.[0-9].*))$", RegexOptions.Compiled);
    private readonly string _debugLogPath;
    private DocumentProperties? _documentProperties;
    private readonly List<(int AfterParagraphIndex, List<ParagraphModel> Paragraphs)> _pendingRecoveredParagraphInsertions = new();
    private sealed class RecoveredCell
    {
        public string Text { get; set; } = string.Empty;
        public ParagraphModel? SourceParagraph { get; set; }
        public List<RunModel> SourceRuns { get; } = new();
    }

    private sealed class CompactGridParseResult
    {
        public int BaseGap { get; set; }
        public int SlotCount { get; set; }
        public int ColumnCount { get; set; }
        public List<Dictionary<int, RecoveredCell>> Rows { get; } = new();
        public RecoveredCell? TitleCell { get; set; }
        public List<ParagraphModel> TrailingParagraphs { get; } = new();
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
        // Generate unique log file name to avoid conflicts in multi-threaded scenarios
        _debugLogPath = Path.Combine(Path.GetTempPath(), $"{_debugLogBaseName}_{Guid.NewGuid():N}.log");
    }

    public void ParseTables(DocumentModel document)
    {
        _documentProperties = document.Properties;
        _pendingRecoveredParagraphInsertions.Clear();
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
            activeCtx.SourceParagraphs.Add(para);

            // Extract TAP properties for table alignment and cell widths
            TapBase? tapForParagraph = ResolveParagraphTap(para);
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

                props.Floating = MergeFloatingTableProperties(props.Floating, tapForParagraph.Floating);

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

            if (TryConsumeCompactTableParagraph(para, activeCtx, tapForParagraph))
            {
                continue;
            }

            // Cell boundary detection. Row end is a cell with ONLY \x07.
            string textContent = string.Join("", para.Runs.Select(r => r.Text));
            bool isRowEnd = string.IsNullOrWhiteSpace(textContent.Replace("\x07", ""));

            if (isRowEnd)
            {
                if (tapForParagraph != null &&
                    (activeCtx.CurrentCellParagraphs.Count > 0 || activeCtx.CellsInCurrentRow.Count > 0))
                {
                    activeCtx.CurrentRowTap = MergeDeferredRowTap(activeCtx.CurrentRowTap, tapForParagraph);
                }

                if (activeCtx.CurrentCellParagraphs.Count == 0 &&
                    activeCtx.CellsInCurrentRow.Count == 0 &&
                    tapForParagraph != null &&
                    activeCtx.Table.Rows.Count > 0)
                {
                    ApplyDeferredRowTap(activeCtx, tapForParagraph);
                }

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

                int trailingBoundaryCount = GetTrailingCellBoundaryColumnCount(para);
                if (trailingBoundaryCount == 0 && textContent.Contains('\x07'))
                {
                    trailingBoundaryCount = 1;
                }

                if (trailingBoundaryCount > 1 &&
                    activeCtx.CurrentCellParagraphs.Count > 0 &&
                    activeCtx.CurrentCellParagraphs.All(p => p.Type == ParagraphType.NestedTable && p.NestedTable != null))
                {
                    // Some mixed parent/nested rows encode the first boundary for the
                    // nested-table host cell and the second one for the following
                    // sibling cell on the same paragraph. Close the nested cell first
                    // so the trailing text lands in the next cell instead of being
                    // swallowed into the nested-table cell.
                    FlushCurrentCell(activeCtx);
                    trailingBoundaryCount--;
                }

                activeCtx.CurrentCellParagraphs.Add(cellParagraph);

                for (int boundaryIndex = 0; boundaryIndex < trailingBoundaryCount; boundaryIndex++)
                {
                    // A paragraph can carry more than one trailing cell marker when a
                    // row ends with additional blank cells.
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

        if (_pendingRecoveredParagraphInsertions.Count > 0)
        {
            ApplyPendingRecoveredParagraphInsertions(document);
        }

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

    private void ApplyPendingRecoveredParagraphInsertions(DocumentModel document)
    {
        int insertedCount = 0;
        foreach (var pendingInsertion in _pendingRecoveredParagraphInsertions
            .OrderBy(insertion => insertion.AfterParagraphIndex)
            .ThenBy(insertion => insertion.Paragraphs.Count))
        {
            int insertIndex = Math.Clamp(pendingInsertion.AfterParagraphIndex + 1 + insertedCount, 0, document.Paragraphs.Count);
            document.Paragraphs.InsertRange(insertIndex, pendingInsertion.Paragraphs);
            insertedCount += pendingInsertion.Paragraphs.Count;
        }

        for (int index = 0; index < document.Paragraphs.Count; index++)
        {
            document.Paragraphs[index].Index = index;
        }
    }

    private static void ApplyFallbackHeaderShading(DocumentModel document)
    {
        foreach (var table in document.Tables)
        {
            ApplyFallbackHeaderShading(table, document.Theme);
        }
    }

    private static void ApplyFallbackHeaderShading(TableModel table, ThemeModel theme)
    {
        foreach (var row in table.Rows)
        {
            if (!NeedsFallbackHeaderShading(row))
                continue;

            var fallbackShading = CreateFallbackHeaderShading(row, table, theme);
            if (fallbackShading == null)
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
                        ApplyFallbackHeaderShading(paragraph.NestedTable, theme);
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

    private static ShadingInfo? CreateFallbackHeaderShading(TableRowModel row, TableModel table, ThemeModel theme)
    {
        int backgroundColor = ResolveFallbackHeaderBackgroundColor(row, table, theme);

        return new ShadingInfo
        {
            Pattern = ShadingPattern.Clear,
            PatternVal = "clear",
            BackgroundColor = backgroundColor
        };
    }

    private static int ResolveFallbackHeaderBackgroundColor(TableRowModel row, TableModel table, ThemeModel theme)
    {
        var borderColor = FindDominantBorderColor(row, table);
        if (borderColor != 0)
            return borderColor;

        // Fall back to the theme accent only when the table itself does not expose
        // a stronger color signal in its borders.
        return theme.ColorMap.ContainsKey("accent1")
            ? 0x01000004
            : 0x00BD814F;
    }

    private static int FindDominantBorderColor(TableRowModel row, TableModel table)
    {
        var candidates = row.Cells
            .SelectMany(GetBorderCandidates)
            .Concat(GetBorderCandidates(table.Properties))
            .Where(color => color != 0)
            .GroupBy(color => color)
            .OrderByDescending(group => group.Count())
            .ThenByDescending(group => group.Key)
            .Select(group => group.Key)
            .FirstOrDefault();

        return candidates;
    }

    private static IEnumerable<int> GetBorderCandidates(TableCellModel cell)
    {
        if (cell.Properties == null)
            yield break;

        foreach (var color in GetBorderCandidates(cell.Properties))
            yield return color;
    }

    private static IEnumerable<int> GetBorderCandidates(TableCellProperties? properties)
    {
        if (properties == null)
            yield break;

        if (properties.BorderTop?.Color is int top && top != 0)
            yield return top;
        if (properties.BorderBottom?.Color is int bottom && bottom != 0)
            yield return bottom;
        if (properties.BorderLeft?.Color is int left && left != 0)
            yield return left;
        if (properties.BorderRight?.Color is int right && right != 0)
            yield return right;
    }

    private static IEnumerable<int> GetBorderCandidates(TableProperties? properties)
    {
        if (properties == null)
            yield break;

        if (properties.BorderTop?.Color is int top && top != 0)
            yield return top;
        if (properties.BorderBottom?.Color is int bottom && bottom != 0)
            yield return bottom;
        if (properties.BorderLeft?.Color is int left && left != 0)
            yield return left;
        if (properties.BorderRight?.Color is int right && right != 0)
            yield return right;
        if (properties.BorderInsideH?.Color is int insideH && insideH != 0)
            yield return insideH;
        if (properties.BorderInsideV?.Color is int insideV && insideV != 0)
            yield return insideV;
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
            if (!looksLikeNestedSectionTitle)
                continue;

            var placeholderChild = TryBuildNestedPlaceholderTable(document, paragraphIndex, nextSectionTitleIndex);

            if (placeholderChild == null)
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

    private void AppendFollowingMetadataRows(DocumentModel document, TableModel table)
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
        var paragraphList = paragraphs.ToList();

        int tapColumnCount = paragraphList
            .Select(GetTapColumnCount)
            .Where(count => count > 0)
            .DefaultIfEmpty(0)
            .Max();

        int markerHintColumnCount = paragraphList
            .Select(GetMarkerHintColumnCount)
            .Where(count => count > 0)
            .DefaultIfEmpty(0)
            .Max();

        tapColumnCount = Math.Max(tapColumnCount, markerHintColumnCount);

        if (tapColumnCount >= 2)
            return tapColumnCount;

        var rowCandidates = paragraphList
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

    private static int GetMarkerHintColumnCount(ParagraphModel paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.RawText))
            return 0;

        int maxMarkerRun = MultiCellBoundaryRegex.Matches(paragraph.RawText)
            .Cast<Match>()
            .Select(match => match.Length)
            .DefaultIfEmpty(0)
            .Max();

        return maxMarkerRun >= 2 ? InferPlaceholderColumnCount(maxMarkerRun) : 0;
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

    private TableRowModel BuildRecoveredRow(List<RecoveredCell> cellTexts, int rowIndex, int columnCount)
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

            ApplyRecoveredPapOverrides(paragraph, recoveredCell);

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

        if (TryConsumeStructuredCompactGridParagraph(paragraph, ctx, tapForParagraph))
            return true;

        var rowCandidates = BuildRecoveredCellRows(paragraph, GetRowCandidates(paragraph).ToList());
        if (rowCandidates.Count == 0 &&
            !string.IsNullOrEmpty(paragraph.RawText) &&
            paragraph.RawText.All(ch => ch == '\x07'))
        {
            int markerOnlyColumnCount = tapForParagraph?.CellWidths?.Length ?? GetMarkerHintColumnCount(paragraph);
            if (markerOnlyColumnCount >= 2)
            {
                rowCandidates = BuildMarkerOnlyRecoveredRows(paragraph, paragraph.RawText.Length, markerOnlyColumnCount);
            }
        }

        if (rowCandidates.Count == 0)
            return false;

        bool encodesMultipleCells = rowCandidates.Count > 1 || rowCandidates.Any(cells => cells.Count > 1);
        if (!encodesMultipleCells)
            return false;

        int inferredCompactColumnCount = 0;
        if (rowCandidates.All(cells => cells.Count == 1))
        {
            inferredCompactColumnCount = InferCompactTrailingEmptyColumnCount(paragraph);
        }

        var recoveredRowAlignments = GetCompactRowAlignments(paragraph, rowCandidates.Count);
        var recoveredRowTaps = GetCompactRowTaps(paragraph, rowCandidates.Count, tapForParagraph);
        int compactColumnCount = Math.Max(
            inferredCompactColumnCount,
            rowCandidates.Select(cells => cells.Count).DefaultIfEmpty(0).Max());
        bool canUseDocumentWidthFallback = rowCandidates.Count > 1;
        var widthTemplate = GetWidthTemplate(recoveredRowTaps, compactColumnCount)
            ?? (canUseDocumentWidthFallback ? BuildCompactWidthTemplate(tapForParagraph, compactColumnCount, _documentProperties) : null);
        if (widthTemplate != null && widthTemplate.Length > 0)
        {
            foreach (var rowTap in recoveredRowTaps)
            {
                ApplyCompactGridWidthTemplate(rowTap, widthTemplate);
            }

            ctx.Table.Properties ??= new TableProperties();
            if (ctx.Table.Properties.PreferredWidth <= compactColumnCount)
            {
                ctx.Table.Properties.PreferredWidth = widthTemplate.Sum();
            }
        }

        if (tapForParagraph?.RowHeight > 0 && rowCandidates.Count > 0)
        {
            while (recoveredRowTaps.Count < rowCandidates.Count)
            {
                recoveredRowTaps.Add(CloneTapBase(recoveredRowTaps.LastOrDefault()) ?? CloneTapBase(tapForParagraph));
            }

            int lastRecoveredRowIndex = rowCandidates.Count - 1;
            recoveredRowTaps[lastRecoveredRowIndex] = MergeDeferredRowTap(recoveredRowTaps[lastRecoveredRowIndex], tapForParagraph);
        }

        while (TryExtractTrailingCompactRowParagraph(rowCandidates, compactColumnCount, out var trailingCell))
        {
            _pendingRecoveredParagraphInsertions.Add((paragraph.Index, new List<ParagraphModel> { BuildRecoveredParagraph(trailingCell) }));
            if (recoveredRowTaps.Count > rowCandidates.Count)
            {
                recoveredRowTaps.RemoveAt(recoveredRowTaps.Count - 1);
            }
            if (recoveredRowAlignments.Count > rowCandidates.Count)
            {
                recoveredRowAlignments.RemoveAt(recoveredRowAlignments.Count - 1);
            }
        }

        int recoveredRowIndex = 0;

        foreach (var recoveredCells in rowCandidates)
        {
            ctx.CurrentRowTap = recoveredRowIndex < recoveredRowTaps.Count
                ? recoveredRowTaps[recoveredRowIndex]
                : tapForParagraph;

            while (inferredCompactColumnCount > 0 && recoveredCells.Count < inferredCompactColumnCount)
            {
                recoveredCells.Add(new RecoveredCell { SourceParagraph = paragraph });
            }

            foreach (var recoveredCell in recoveredCells)
            {
                var recoveredParagraph = BuildRecoveredParagraph(recoveredCell);
                if (recoveredRowIndex < recoveredRowAlignments.Count)
                {
                    recoveredParagraph.Properties ??= new ParagraphProperties();
                    recoveredParagraph.Properties.Alignment = recoveredRowAlignments[recoveredRowIndex];
                }

                ctx.CurrentCellParagraphs.Add(recoveredParagraph);
                FlushCurrentCell(ctx);
            }

            FlushCurrentRow(ctx);
            recoveredRowIndex++;
        }

        return true;
    }

    private static List<List<RecoveredCell>> BuildMarkerOnlyRecoveredRows(ParagraphModel paragraph, int markerCount, int columnCount)
    {
        var rows = new List<List<RecoveredCell>>();
        if (markerCount <= 0 || columnCount <= 0)
            return rows;

        int remainder = markerCount % columnCount;
        if (remainder > 0 && markerCount > columnCount)
        {
            rows.Add(BuildMarkerOnlyRecoveredRow(paragraph, remainder));
        }

        int remainingMarkers = markerCount - (rows.Count > 0 ? remainder : 0);
        while (remainingMarkers > 0)
        {
            int cellCount = Math.Min(columnCount, remainingMarkers);
            rows.Add(BuildMarkerOnlyRecoveredRow(paragraph, cellCount));
            remainingMarkers -= cellCount;
        }

        return rows;
    }

    private static List<RecoveredCell> BuildMarkerOnlyRecoveredRow(ParagraphModel paragraph, int cellCount)
    {
        var row = new List<RecoveredCell>(cellCount);
        for (int cellIndex = 0; cellIndex < cellCount; cellIndex++)
        {
            row.Add(new RecoveredCell { SourceParagraph = paragraph });
        }

        return row;
    }

    private static bool TryExtractTrailingCompactRowParagraph(List<List<RecoveredCell>> rowCandidates, int columnCount, out RecoveredCell trailingCell)
    {
        trailingCell = new RecoveredCell();
        if (rowCandidates.Count == 0 || columnCount <= 1)
            return false;

        var lastRow = rowCandidates[^1];
        if (lastRow.Count != 1 || string.IsNullOrWhiteSpace(lastRow[0].Text))
            return false;

        if (rowCandidates.Count < 2)
            return false;

        var previousRow = rowCandidates[^2];
        if (previousRow.Count < columnCount)
            return false;

        trailingCell = lastRow[0];
        rowCandidates.RemoveAt(rowCandidates.Count - 1);
        return true;
    }

    private bool TryConsumeStructuredCompactGridParagraph(ParagraphModel paragraph, TableContext ctx, TapBase? tapForParagraph)
    {
        if (!TryParseCompactGridParagraph(paragraph, out var grid))
            return false;

        var rowAlignments = GetCompactRowAlignments(paragraph, grid.Rows.Count + (grid.TitleCell != null ? 1 : 0));
        var rowTaps = GetCompactRowTaps(paragraph, grid.Rows.Count + (grid.TitleCell != null ? 1 : 0), tapForParagraph);
        var widthTemplate = BuildCompactGridWidthTemplate(tapForParagraph, grid.ColumnCount, grid.BaseGap, grid.SlotCount);

        // Compact-grid tables are synthesized from a flattened paragraph stream.
        // Paragraph-level TAP borders/shading are often row-scoped formatting, not
        // reliable whole-table properties, so clear any speculative table-level
        // border/shading that may have been pre-populated before compact parsing.
        if (ctx.Table.Properties != null)
        {
            ctx.Table.Properties.BorderTop = null;
            ctx.Table.Properties.BorderBottom = null;
            ctx.Table.Properties.BorderLeft = null;
            ctx.Table.Properties.BorderRight = null;
            ctx.Table.Properties.BorderInsideH = null;
            ctx.Table.Properties.BorderInsideV = null;
            ctx.Table.Properties.Shading = null;
        }

        int targetRowIndex = 0;
        if (grid.TitleCell != null)
        {
            var titleTap = targetRowIndex < rowTaps.Count ? CloneTapBase(rowTaps[targetRowIndex]) : CloneTapBase(tapForParagraph);
            ApplyCompactGridWidthTemplate(titleTap, widthTemplate);
            ApplyCompactGridTableProperties(ctx.Table, titleTap);
            ctx.Table.Rows.Add(BuildCompactGridTitleRow(grid.TitleCell, targetRowIndex, grid.ColumnCount, titleTap, rowAlignments.ElementAtOrDefault(targetRowIndex)));
            ctx.RowTaps.Add(titleTap);
            targetRowIndex++;
        }

        foreach (var rowCells in grid.Rows)
        {
            var rowTap = targetRowIndex < rowTaps.Count ? CloneTapBase(rowTaps[targetRowIndex]) : CloneTapBase(tapForParagraph);
            ApplyCompactGridWidthTemplate(rowTap, widthTemplate);
            ApplyCompactGridTableProperties(ctx.Table, rowTap);
            ctx.Table.Rows.Add(BuildCompactGridRow(rowCells, targetRowIndex, grid.ColumnCount, rowTap, rowAlignments.ElementAtOrDefault(targetRowIndex)));
            ctx.RowTaps.Add(rowTap);
            targetRowIndex++;
        }

        ctx.RowIndex = ctx.Table.Rows.Count;

        if (widthTemplate.Length > 0)
        {
            ctx.Table.Properties ??= new TableProperties();
            if (ctx.Table.Properties.PreferredWidth <= 0)
            {
                ctx.Table.Properties.PreferredWidth = widthTemplate.Sum();
            }
        }

        if (grid.TrailingParagraphs.Count > 0)
        {
            _pendingRecoveredParagraphInsertions.Add((paragraph.Index, grid.TrailingParagraphs));
        }

        return true;
    }

    private bool TryParseCompactGridParagraph(ParagraphModel paragraph, out CompactGridParseResult result)
    {
        result = new CompactGridParseResult();

        if (string.IsNullOrEmpty(paragraph.RawText))
            return false;

        var tokens = CellBoundaryRegex.Split(paragraph.RawText)
            .Where(token => !string.IsNullOrWhiteSpace(token))
            .ToList();
        var gaps = CellBoundaryRegex.Matches(paragraph.RawText)
            .Cast<Match>()
            .Select(match => match.Length)
            .ToList();

        if (tokens.Count < 4 || gaps.Count < tokens.Count - 1)
            return false;

        int baseGap = gaps.Where(gap => gap > 0).DefaultIfEmpty(0).Min();
        if (baseGap <= 1)
            return false;

        var recoveredCells = BuildRecoveredCells(paragraph, tokens);
        int gridTokenStart = 0;
        if (gaps[0] != baseGap)
        {
            result.TitleCell = recoveredCells[0];
            gridTokenStart = 1;
        }

        int slotCount = 1;
        for (int gapIndex = gridTokenStart; gapIndex < gaps.Count; gapIndex++)
        {
            if (gaps[gapIndex] != baseGap)
                break;

            slotCount++;
        }

        if (slotCount < 2)
            return false;

        result.BaseGap = baseGap;
        result.SlotCount = slotCount;
        result.ColumnCount = ((slotCount - 1) * baseGap) + 1;

        int slotIndex = 0;
        AddCompactGridPlacement(result.Rows, 0, 0, recoveredCells[gridTokenStart]);

        for (int tokenIndex = gridTokenStart + 1; tokenIndex < recoveredCells.Count; tokenIndex++)
        {
            int gap = gaps[tokenIndex - 1];
            if (gap % baseGap != 0)
                return false;

            slotIndex += gap / baseGap;
            int rowIndex = slotIndex / slotCount;
            int columnIndex = (slotIndex % slotCount) * baseGap;
            AddCompactGridPlacement(result.Rows, rowIndex, columnIndex, recoveredCells[tokenIndex]);
        }

        if (result.Rows.Count == 0)
            return false;

        while (TryExtractTrailingCompactGridParagraph(result.Rows, result.ColumnCount, out var trailingCell))
        {
            result.TrailingParagraphs.Insert(0, BuildRecoveredParagraph(trailingCell));
        }

        return result.Rows.Count > 0;
    }

    private static void AddCompactGridPlacement(List<Dictionary<int, RecoveredCell>> rows, int rowIndex, int columnIndex, RecoveredCell recoveredCell)
    {
        while (rows.Count <= rowIndex)
        {
            rows.Add(new Dictionary<int, RecoveredCell>());
        }

        rows[rowIndex][columnIndex] = recoveredCell;
    }

    private static bool TryExtractTrailingCompactGridParagraph(List<Dictionary<int, RecoveredCell>> rows, int columnCount, out RecoveredCell trailingCell)
    {
        trailingCell = new RecoveredCell();
        if (rows.Count == 0)
            return false;

        var lastRow = rows[^1];
        if (lastRow.Count != 1)
            return false;

        if (!lastRow.TryGetValue(0, out var candidate) || string.IsNullOrWhiteSpace(candidate.Text))
            return false;

        bool precedingRowHasMultipleSignals = rows.Count >= 2 && rows[^2].Count > 0;
        bool rowStartsNewLogicalBlock = rows.Count >= 2 && rows[^2].Keys.Max() < columnCount - 1;
        if (!precedingRowHasMultipleSignals || !rowStartsNewLogicalBlock)
            return false;

        trailingCell = candidate;
        rows.RemoveAt(rows.Count - 1);
        return true;
    }

    private static int[] BuildCompactGridWidthTemplate(TapBase? tap, int columnCount, int baseGap, int slotCount)
    {
        if (tap?.CellWidths != null && tap.CellWidths.Length == columnCount)
            return tap.CellWidths.ToArray();

        if (tap?.CellWidths != null && tap.CellWidths.Length == columnCount + 2)
            return tap.CellWidths.Skip(1).Take(columnCount).ToArray();

        int preferredWidth = tap?.TableWidth ?? 0;
        if (preferredWidth <= 0)
        {
            preferredWidth = 9360;
        }

        int spacerColumnCount = columnCount - slotCount;
        int spacerWidth = spacerColumnCount > 0
            ? Math.Max(1, preferredWidth / Math.Max(1, (slotCount * 4) + spacerColumnCount))
            : 0;
        int contentWidth = Math.Max(1, (preferredWidth - (spacerWidth * spacerColumnCount)) / Math.Max(1, slotCount));

        var widths = new int[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            widths[columnIndex] = columnIndex % baseGap == 0 ? contentWidth : spacerWidth;
        }

        return widths;
    }

    private static int[]? BuildCompactWidthTemplate(TapBase? tap, int columnCount, DocumentProperties? documentProperties)
    {
        if (columnCount <= 0)
            return null;

        if (tap?.CellWidths != null && tap.CellWidths.Length == columnCount)
            return tap.CellWidths.ToArray();

        if (tap?.CellWidths != null && tap.CellWidths.Length == columnCount + 2)
            return tap.CellWidths.Skip(1).Take(columnCount).ToArray();

        int preferredWidth = tap?.TableWidth ?? 0;
        if (preferredWidth > columnCount)
            return DistributeWidthAcrossColumns(preferredWidth, columnCount);

        int documentColumnWidth = documentProperties?.DxaColumns ?? 0;
        if (documentColumnWidth > 0)
        {
            var widths = new int[columnCount];
            Array.Fill(widths, documentColumnWidth);
            return widths;
        }

        int contentWidth = 0;
        if (documentProperties != null)
        {
            contentWidth = Math.Max(0, documentProperties.PageWidth - documentProperties.MarginLeft - documentProperties.MarginRight);
        }

        int cellSpacing = tap?.CellSpacing != 0
            ? tap.CellSpacing
            : (tap?.GapHalf ?? 0) * 2;

        if (contentWidth > 0 && cellSpacing > 0)
        {
            contentWidth += cellSpacing;
        }

        return contentWidth > 0
            ? DistributeWidthAcrossColumns(contentWidth, columnCount)
            : null;
    }

    private static int[] DistributeWidthAcrossColumns(int totalWidth, int columnCount)
    {
        int baseWidth = totalWidth / columnCount;
        int remainder = totalWidth % columnCount;
        var widths = new int[columnCount];
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            widths[columnIndex] = baseWidth + (columnIndex < remainder ? 1 : 0);
        }

        return widths;
    }

    private static void ApplyCompactGridWidthTemplate(TapBase? tap, int[] widthTemplate)
    {
        if (tap == null || widthTemplate.Length == 0)
            return;

        tap.CellWidths = widthTemplate.ToArray();
        if (tap.TableWidth <= 0)
        {
            tap.TableWidth = widthTemplate.Sum();
        }
    }

    private TableRowModel BuildCompactGridTitleRow(RecoveredCell titleCell, int rowIndex, int columnCount, TapBase? rowTap, ParagraphAlignment alignment)
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

        var paragraph = BuildRecoveredParagraph(titleCell);
        paragraph.Properties ??= new ParagraphProperties();
        paragraph.Properties.Alignment = alignment;

        var cell = new TableCellModel
        {
            Index = 0,
            RowIndex = rowIndex,
            ColumnIndex = 0,
            ColumnSpan = columnCount,
            Paragraphs = new List<ParagraphModel> { paragraph },
            Properties = new TableCellProperties()
        };

        ConfigureCompactGridCell(cell, rowTap, 0, columnCount);
        row.Cells.Add(cell);
        return row;
    }

    private TableRowModel BuildCompactGridRow(Dictionary<int, RecoveredCell> rowCells, int rowIndex, int columnCount, TapBase? rowTap, ParagraphAlignment alignment)
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

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            ParagraphModel paragraph;
            if (rowCells.TryGetValue(columnIndex, out var recoveredCell))
            {
                paragraph = BuildRecoveredParagraph(recoveredCell);
            }
            else
            {
                paragraph = new ParagraphModel
                {
                    Type = ParagraphType.Normal,
                    Properties = new ParagraphProperties(),
                    Runs = new List<RunModel>()
                };
            }

            paragraph.Properties ??= new ParagraphProperties();
            paragraph.Properties.Alignment = alignment;

            var cell = new TableCellModel
            {
                Index = columnIndex,
                RowIndex = rowIndex,
                ColumnIndex = columnIndex,
                Paragraphs = new List<ParagraphModel> { paragraph },
                Properties = new TableCellProperties()
            };

            ConfigureCompactGridCell(cell, rowTap, columnIndex, 1);
            row.Cells.Add(cell);
        }

        return row;
    }

    private static void ConfigureCompactGridCell(TableCellModel cell, TapBase? rowTap, int columnIndex, int columnSpan)
    {
        cell.ColumnSpan = Math.Max(1, columnSpan);
        cell.Properties ??= new TableCellProperties();

        if (rowTap?.CellWidths != null && rowTap.CellWidths.Length > columnIndex)
        {
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

        if (rowTap?.CellBorders != null && rowTap.CellBorders.Length > columnIndex)
        {
            var cellBorders = rowTap.CellBorders[columnIndex];
            if (cellBorders != null)
            {
                cell.Properties.BorderTop = cellBorders.Top;
                cell.Properties.BorderBottom = cellBorders.Bottom;
                cell.Properties.BorderLeft = cellBorders.Left;
                cell.Properties.BorderRight = cellBorders.Right;
            }
        }

        if (rowTap?.CellShadings != null && rowTap.CellShadings.Length > columnIndex)
        {
            var shading = rowTap.CellShadings[columnIndex];
            if (shading != null)
            {
                cell.Properties.Shading = shading;
            }
        }

        if (rowTap?.CellVerticalAlignments != null && rowTap.CellVerticalAlignments.Length > columnIndex)
        {
            cell.Properties.VerticalAlignment = (VerticalAlignment)rowTap.CellVerticalAlignments[columnIndex];
        }
    }

    private static void ApplyCompactGridTableProperties(TableModel table, TapBase? rowTap)
    {
        if (rowTap == null)
            return;

        table.Properties ??= new TableProperties();

        if (table.Properties.Alignment == TableAlignment.Left && rowTap.Justification != 0)
        {
            table.Properties.Alignment = rowTap.Justification switch
            {
                1 => TableAlignment.Center,
                2 => TableAlignment.Right,
                _ => TableAlignment.Left
            };
        }

        if (table.Properties.CellSpacing == 0)
        {
            table.Properties.CellSpacing = rowTap.CellSpacing != 0
                ? rowTap.CellSpacing
                : (rowTap.GapHalf != 0 ? rowTap.GapHalf * 2 : 0);
        }

        if (table.Properties.Indent == 0 && rowTap.IndentLeft != 0)
        {
            table.Properties.Indent = rowTap.IndentLeft;
        }

        if (table.Properties.PreferredWidth == 0 && rowTap.TableWidth != 0)
        {
            table.Properties.PreferredWidth = rowTap.TableWidth;
        }

        table.Properties.Floating = MergeFloatingTableProperties(table.Properties.Floating, rowTap.Floating);

    }

    private List<TapBase?> GetCompactRowTaps(ParagraphModel paragraph, int expectedRowCount, TapBase? fallbackTap)
    {
        var rowTaps = new List<TapBase?>();
        var firstRun = paragraph.Runs.FirstOrDefault();
        var lastRun = paragraph.Runs.LastOrDefault();
        if (firstRun == null || lastRun == null || expectedRowCount <= 0)
        {
            FillMissingCompactRowTaps(rowTaps, expectedRowCount, fallbackTap);
            return rowTaps;
        }

        int startCp = firstRun.CharacterPosition;
        int endCp = lastRun.CharacterPosition + Math.Max(1, lastRun.CharacterLength);
        TapBase? previousTap = null;

        for (int cp = startCp; cp <= endCp; cp++)
        {
            var currentTap = _fkpParser.GetPapAtCp(cp)?.Tap;
            if (AreEquivalentCompactRowTaps(previousTap, currentTap))
                continue;

            previousTap = CloneTapBase(currentTap);
            rowTaps.Add(CloneTapBase(currentTap));
        }

        FillMissingCompactRowTaps(rowTaps, expectedRowCount, fallbackTap);
        return rowTaps;
    }

    private static void FillMissingCompactRowTaps(List<TapBase?> rowTaps, int expectedRowCount, TapBase? fallbackTap)
    {
        if (rowTaps.Count == 0)
        {
            rowTaps.Add(CloneTapBase(fallbackTap));
        }

        while (rowTaps.Count < expectedRowCount)
        {
            rowTaps.Add(CloneTapBase(rowTaps.LastOrDefault()) ?? CloneTapBase(fallbackTap));
        }
    }

    private static bool AreEquivalentCompactRowTaps(TapBase? left, TapBase? right)
    {
        if (left == null && right == null)
            return true;

        if (left == null || right == null)
            return false;

        return left.RowHeight == right.RowHeight &&
               left.HeightIsExact == right.HeightIsExact &&
               left.CantSplit == right.CantSplit &&
               left.IsHeaderRow == right.IsHeaderRow &&
               left.Justification == right.Justification &&
               left.TableWidth == right.TableWidth &&
               left.IndentLeft == right.IndentLeft &&
               left.CellSpacing == right.CellSpacing &&
             FloatingEqual(left.Floating, right.Floating) &&
               BordersEqual(left.BorderTop, right.BorderTop) &&
               BordersEqual(left.BorderBottom, right.BorderBottom) &&
               BordersEqual(left.BorderLeft, right.BorderLeft) &&
               BordersEqual(left.BorderRight, right.BorderRight) &&
               BordersEqual(left.BorderInsideH, right.BorderInsideH) &&
               BordersEqual(left.BorderInsideV, right.BorderInsideV) &&
               CellBordersEqual(left.CellBorders, right.CellBorders) &&
               CellMergesEqual(left.CellMerges, right.CellMerges) &&
               CellShadingsEqual(left.CellShadings, right.CellShadings) &&
               ((left.CellVerticalAlignments == null && right.CellVerticalAlignments == null) ||
                (left.CellVerticalAlignments != null && right.CellVerticalAlignments != null && left.CellVerticalAlignments.SequenceEqual(right.CellVerticalAlignments))) &&
               ((left.CellWidths == null && right.CellWidths == null) ||
                (left.CellWidths != null && right.CellWidths != null && left.CellWidths.SequenceEqual(right.CellWidths)));
    }

    private static bool BordersEqual(BorderInfo? left, BorderInfo? right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left == null || right == null)
            return false;

        return left.Style == right.Style &&
               left.Width == right.Width &&
               left.Space == right.Space &&
               left.Color == right.Color;
    }

    private static bool FloatingEqual(TableFloatingProperties? left, TableFloatingProperties? right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left == null || right == null)
            return false;

        return left.HorizontalPosition == right.HorizontalPosition &&
               left.VerticalPosition == right.VerticalPosition &&
               left.LeftFromText == right.LeftFromText &&
               left.RightFromText == right.RightFromText &&
               left.TopFromText == right.TopFromText &&
               left.BottomFromText == right.BottomFromText &&
               left.AllowOverlap == right.AllowOverlap;
    }

    private static bool CellBordersEqual(CellBorderInfo?[]? left, CellBorderInfo?[]? right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left == null || right == null || left.Length != right.Length)
            return false;

        for (int i = 0; i < left.Length; i++)
        {
            var leftBorder = left[i];
            var rightBorder = right[i];

            if (ReferenceEquals(leftBorder, rightBorder))
                continue;

            if (leftBorder == null || rightBorder == null)
                return false;

            if (!BordersEqual(leftBorder.Top, rightBorder.Top) ||
                !BordersEqual(leftBorder.Bottom, rightBorder.Bottom) ||
                !BordersEqual(leftBorder.Left, rightBorder.Left) ||
                !BordersEqual(leftBorder.Right, rightBorder.Right))
            {
                return false;
            }
        }

        return true;
    }

    private static bool CellMergesEqual(CellMergeFlags[]? left, CellMergeFlags[]? right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left == null || right == null || left.Length != right.Length)
            return false;

        for (int i = 0; i < left.Length; i++)
        {
            if (left[i].HorizFirst != right[i].HorizFirst ||
                left[i].HorizMerged != right[i].HorizMerged ||
                left[i].VertFirst != right[i].VertFirst ||
                left[i].VertMerged != right[i].VertMerged)
            {
                return false;
            }
        }

        return true;
    }

    private static bool CellShadingsEqual(ShadingInfo?[]? left, ShadingInfo?[]? right)
    {
        if (ReferenceEquals(left, right))
            return true;

        if (left == null || right == null || left.Length != right.Length)
            return false;

        for (int i = 0; i < left.Length; i++)
        {
            var leftShading = left[i];
            var rightShading = right[i];

            if (ReferenceEquals(leftShading, rightShading))
                continue;

            if (leftShading == null || rightShading == null)
                return false;

            if (leftShading.Pattern != rightShading.Pattern ||
                leftShading.PatternVal != rightShading.PatternVal ||
                leftShading.ForegroundColor != rightShading.ForegroundColor ||
                leftShading.BackgroundColor != rightShading.BackgroundColor)
            {
                return false;
            }
        }

        return true;
    }

    private static TapBase? CloneTapBase(TapBase? source)
    {
        if (source == null)
            return null;

        return new TapBase
        {
            RowHeight = source.RowHeight,
            HeightIsExact = source.HeightIsExact,
            Justification = source.Justification,
            IsHeaderRow = source.IsHeaderRow,
            CellSpacing = source.CellSpacing,
            TableWidth = source.TableWidth,
            IndentLeft = source.IndentLeft,
            Floating = CloneFloatingTableProperties(source.Floating),
            GapHalf = source.GapHalf,
            CellWidths = source.CellWidths?.ToArray(),
            CantSplit = source.CantSplit,
            CellMerges = source.CellMerges?.Select(flags => new CellMergeFlags
            {
                HorizFirst = flags.HorizFirst,
                HorizMerged = flags.HorizMerged,
                VertFirst = flags.VertFirst,
                VertMerged = flags.VertMerged
            }).ToArray(),
            BorderTop = source.BorderTop,
            BorderBottom = source.BorderBottom,
            BorderLeft = source.BorderLeft,
            BorderRight = source.BorderRight,
            BorderInsideH = source.BorderInsideH,
            BorderInsideV = source.BorderInsideV,
            Shading = source.Shading,
            CellBorders = source.CellBorders?.ToArray(),
            CellShadings = source.CellShadings?.ToArray(),
            CellVerticalAlignments = source.CellVerticalAlignments?.ToArray()
        };
    }

    private static int InferCompactTrailingEmptyColumnCount(ParagraphModel paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.RawText))
            return 0;

        return MultiCellBoundaryRegex.Matches(paragraph.RawText)
            .Cast<Match>()
            .Select(match => Math.Max(1, match.Length - 1))
            .DefaultIfEmpty(0)
            .Max();
    }

    private List<ParagraphAlignment> GetCompactRowAlignments(ParagraphModel paragraph, int expectedRowCount)
    {
        var alignments = new List<ParagraphAlignment>();
        var segments = new List<byte>();
        var firstRun = paragraph.Runs.FirstOrDefault();
        var lastRun = paragraph.Runs.LastOrDefault();
        if (firstRun == null || lastRun == null || expectedRowCount <= 0)
            return alignments;

        int startCp = firstRun.CharacterPosition;
        int endCp = lastRun.CharacterPosition + Math.Max(1, lastRun.CharacterLength);
        byte? previousJustification = null;

        for (int cp = startCp; cp <= endCp; cp++)
        {
            var pap = _fkpParser.GetPapAtCp(cp);
            byte justification = pap?.Justification ?? 0;
            if (previousJustification == justification)
                continue;

            previousJustification = justification;
            segments.Add(justification);
        }

        if (segments.Count == 0)
            return alignments;

        alignments.Add(MapParagraphAlignment(segments[0]));
        foreach (var justification in segments.Skip(1))
        {
            if (justification == 0)
                continue;

            alignments.Add(MapParagraphAlignment(justification));
            if (alignments.Count >= expectedRowCount)
                break;
        }

        while (alignments.Count < expectedRowCount)
        {
            alignments.Add(alignments.LastOrDefault());
        }

        return alignments;
    }

    private static ParagraphAlignment MapParagraphAlignment(byte justification)
    {
        return justification switch
        {
            1 => ParagraphAlignment.Center,
            2 => ParagraphAlignment.Right,
            3 => ParagraphAlignment.Justify,
            4 => ParagraphAlignment.Distributed,
            5 => ParagraphAlignment.ThaiJustify,
            _ => ParagraphAlignment.Left
        };
    }

    private ParagraphModel BuildRecoveredParagraph(RecoveredCell recoveredCell)
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

        ApplyRecoveredPapOverrides(paragraph, recoveredCell);

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

    private void ApplyRecoveredPapOverrides(ParagraphModel paragraph, RecoveredCell recoveredCell)
    {
        var firstRun = recoveredCell.SourceRuns.FirstOrDefault();
        if (firstRun == null)
            return;

        var paragraphStartCp = recoveredCell.SourceParagraph?.Runs.FirstOrDefault()?.CharacterPosition ?? firstRun.CharacterPosition;
        var pap = ResolveRecoveredParagraphPap(firstRun.CharacterPosition, paragraphStartCp);
        if (pap == null)
            return;

        paragraph.Properties ??= new ParagraphProperties();

        if (pap.StyleId != 0)
            paragraph.Properties.StyleIndex = pap.StyleId;
        else if (paragraph.Properties.StyleIndex <= 0 && pap.Istd != 0)
            paragraph.Properties.StyleIndex = pap.Istd;

        paragraph.Properties.Alignment = pap.Justification switch
        {
            1 => ParagraphAlignment.Center,
            2 => ParagraphAlignment.Right,
            3 => ParagraphAlignment.Justify,
            4 => ParagraphAlignment.Distributed,
            5 => ParagraphAlignment.ThaiJustify,
            _ => ParagraphAlignment.Left
        };
    }

    private PapBase? ResolveRecoveredParagraphPap(int contentStartCp, int paragraphStartCp)
    {
        var directPap = _fkpParser.GetPapAtCp(contentStartCp);
        if (directPap?.Justification is 1 or 2 or 3 or 4 or 5)
            return directPap;

        int minCp = Math.Max(0, paragraphStartCp);
        int probeStartCp = Math.Max(minCp, contentStartCp - 8);
        for (int cp = contentStartCp - 1; cp >= probeStartCp; cp--)
        {
            var pap = _fkpParser.GetPapAtCp(cp);
            if (pap?.Justification is 1 or 2 or 3 or 4 or 5)
                return pap;
        }

        return directPap;
    }

    private static IEnumerable<List<string>> GetRowCandidates(ParagraphModel paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.RawText))
            yield break;

        foreach (var rawLine in LineBreakRegex.Split(paragraph.RawText))
        {
            bool endsWithCellSeparator = rawLine.EndsWith("\x07", StringComparison.Ordinal);
            var cells = rawLine
                .Split('\x07')
                .Select(NormalizeFlatCellText)
                .ToList();

            int trailingBlankCellsToKeep = endsWithCellSeparator ? 1 : 0;
            while (cells.Count > trailingBlankCellsToKeep && string.IsNullOrWhiteSpace(cells[^1]))
            {
                cells.RemoveAt(cells.Count - 1);
            }

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
        return ResolveParagraphTap(paragraph)?.CellWidths?.Length ?? 0;
    }

    private TapBase? ResolveParagraphTap(ParagraphModel paragraph)
    {
        TapBase? bestRunTap = null;
        int bestRunScore = -1;
        TapBase? bestParagraphTap = null;
        int bestParagraphScore = -1;
        var candidateTaps = new List<TapBase>();

        void ConsiderTap(TapBase? tap, bool preferForGeometry)
        {
            if (tap == null)
                return;

            var clonedTap = CloneTapBase(tap);
            if (clonedTap == null)
                return;

            candidateTaps.Add(clonedTap);

            int score = GetTapScore(tap);
            if (preferForGeometry)
            {
                if (score < bestRunScore)
                    return;

                bestRunTap = CloneTapBase(tap);
                bestRunScore = score;
                return;
            }

            if (score < bestParagraphScore)
                return;

            bestParagraphTap = CloneTapBase(tap);
            bestParagraphScore = score;
        }

        foreach (var run in paragraph.Runs)
        {
            var pap = _fkpParser.GetPapAtCp(run.CharacterPosition);
            ConsiderTap(pap?.Tap, preferForGeometry: true);
        }

        var firstRun = paragraph.Runs.FirstOrDefault();
        int paragraphStartCp = -1;
        int paragraphEndCp = -1;

        if (firstRun != null && !string.IsNullOrEmpty(paragraph.RawText))
        {
            paragraphStartCp = firstRun.CharacterPosition;
            paragraphEndCp = paragraphStartCp + Math.Max(0, paragraph.RawText.Length - 1);
        }
        else if (paragraph.StartCp >= 0)
        {
            paragraphStartCp = paragraph.StartCp;
            paragraphEndCp = paragraph.EndCp >= paragraph.StartCp
                ? paragraph.EndCp
                : paragraph.StartCp + Math.Max(0, paragraph.RawText.Length - 1);
        }

        if (paragraphStartCp >= 0 && paragraphEndCp >= paragraphStartCp)
        {
            bool foundTapInParagraphRange = false;
            for (int cp = paragraphStartCp; cp <= paragraphEndCp; cp++)
            {
                var pap = _fkpParser.GetPapAtCp(cp);
                foundTapInParagraphRange |= pap?.Tap != null;
                ConsiderTap(pap?.Tap, preferForGeometry: false);
            }

            if (!foundTapInParagraphRange)
            {
                int probeEndCp = paragraphEndCp + 64;
                for (int cp = paragraphEndCp + 1; cp <= probeEndCp; cp++)
                {
                    var pap = _fkpParser.GetPapAtCp(cp);
                    ConsiderTap(pap?.Tap, preferForGeometry: false);
                    if (bestParagraphTap != null && bestParagraphScore > 0)
                        break;
                }

                if (bestRunTap == null && bestParagraphTap == null)
                {
                    int probeStartCp = Math.Max(0, paragraphStartCp - 16);
                    for (int cp = paragraphStartCp - 1; cp >= probeStartCp; cp--)
                    {
                        var pap = _fkpParser.GetPapAtCp(cp);
                        ConsiderTap(pap?.Tap, preferForGeometry: false);
                        if (bestParagraphTap != null && bestParagraphScore > 0)
                            break;
                    }
                }
            }
        }

        var bestTap = bestRunTap ?? bestParagraphTap;
        if (bestTap != null && candidateTaps.Count > 1)
        {
            foreach (var candidateTap in candidateTaps)
            {
                bestTap = MergeDeferredRowTap(bestTap, candidateTap);
            }
        }

        return bestTap;
    }

    private static int GetTapScore(TapBase tap)
    {
        int cellCount = tap.CellWidths?.Length ?? 0;
        int borderCount = 0;
        if (tap.BorderTop != null) borderCount++;
        if (tap.BorderBottom != null) borderCount++;
        if (tap.BorderLeft != null) borderCount++;
        if (tap.BorderRight != null) borderCount++;
        if (tap.BorderInsideH != null) borderCount++;
        if (tap.BorderInsideV != null) borderCount++;

        return (cellCount * 1000)
            + Math.Min(Math.Abs(tap.TableWidth), 999)
            + Math.Min(Math.Abs(tap.IndentLeft), 99)
            + (tap.CellSpacing > 0 ? 50 : 0)
            + (tap.GapHalf > 0 ? 25 : 0)
            + borderCount;
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
        var match = SectionTitleRegex.Match(lastCell);
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
            FlipHorizontal = source.FlipHorizontal,
            FlipVertical = source.FlipVertical,
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

    private static void ApplyDeferredRowTap(TableContext ctx, TapBase tap)
    {
        var lastRow = ctx.Table.Rows.LastOrDefault();
        if (lastRow == null)
            return;

        lastRow.Properties ??= new TableRowProperties();
        if (tap.RowHeight > 0)
        {
            lastRow.Properties.Height = tap.RowHeight;
            lastRow.Properties.HeightIsExact = tap.HeightIsExact;
        }

        if (tap.IsHeaderRow)
        {
            lastRow.Properties.IsHeaderRow = true;
        }

        lastRow.Properties.AllowBreakAcrossPages = !tap.CantSplit;
        ApplyTapCellProperties(lastRow, tap);

        if (ctx.RowTaps.Count == 0)
            return;

        ctx.RowTaps[^1] = MergeDeferredRowTap(ctx.RowTaps[^1], tap);
    }

    private static void ApplyTapCellProperties(TableRowModel row, TapBase rowTap)
    {
        foreach (var cell in row.Cells)
        {
            int columnIndex = Math.Max(0, cell.ColumnIndex);
            cell.Properties ??= new TableCellProperties();

            if (rowTap.CellWidths != null && rowTap.CellWidths.Length > columnIndex)
            {
                int width = 0;
                int columnSpan = Math.Max(1, cell.ColumnSpan);
                for (int offset = 0; offset < columnSpan && columnIndex + offset < rowTap.CellWidths.Length; offset++)
                {
                    width += rowTap.CellWidths[columnIndex + offset];
                }

                if (width > 0)
                {
                    cell.Properties.Width = width;
                }
            }

            if (rowTap.CellBorders != null && rowTap.CellBorders.Length > columnIndex)
            {
                var cellBorders = rowTap.CellBorders[columnIndex];
                if (cellBorders != null)
                {
                    cell.Properties.BorderTop = cellBorders.Top;
                    cell.Properties.BorderBottom = cellBorders.Bottom;
                    cell.Properties.BorderLeft = cellBorders.Left;
                    cell.Properties.BorderRight = cellBorders.Right;
                }
            }

            if (rowTap.CellShadings != null && rowTap.CellShadings.Length > columnIndex)
            {
                var shading = rowTap.CellShadings[columnIndex];
                if (shading != null)
                {
                    cell.Properties.Shading = shading;
                }
            }

            if (rowTap.CellVerticalAlignments != null && rowTap.CellVerticalAlignments.Length > columnIndex)
            {
                cell.Properties.VerticalAlignment = (VerticalAlignment)rowTap.CellVerticalAlignments[columnIndex];
            }
        }
    }

    private static TapBase MergeDeferredRowTap(TapBase? existingTap, TapBase deferredTap)
    {
        if (existingTap == null)
            return deferredTap;

        if (existingTap.RowHeight <= 0 && deferredTap.RowHeight > 0)
        {
            existingTap.RowHeight = deferredTap.RowHeight;
            existingTap.HeightIsExact = deferredTap.HeightIsExact;
        }
        else if (deferredTap.RowHeight > 0)
        {
            existingTap.RowHeight = deferredTap.RowHeight;
            existingTap.HeightIsExact = deferredTap.HeightIsExact;
        }

        existingTap.IsHeaderRow |= deferredTap.IsHeaderRow;
        existingTap.CantSplit |= deferredTap.CantSplit;
        existingTap.BorderTop = MergeBorder(existingTap.BorderTop, deferredTap.BorderTop);
        existingTap.BorderBottom = MergeBorder(existingTap.BorderBottom, deferredTap.BorderBottom);
        existingTap.BorderLeft = MergeBorder(existingTap.BorderLeft, deferredTap.BorderLeft);
        existingTap.BorderRight = MergeBorder(existingTap.BorderRight, deferredTap.BorderRight);
        existingTap.BorderInsideH = MergeBorder(existingTap.BorderInsideH, deferredTap.BorderInsideH);
        existingTap.BorderInsideV = MergeBorder(existingTap.BorderInsideV, deferredTap.BorderInsideV);

        if (deferredTap.CellBorders != null)
        {
            if (existingTap.CellBorders == null || existingTap.CellBorders.Length < deferredTap.CellBorders.Length)
            {
                existingTap.CellBorders = deferredTap.CellBorders.Select(CloneCellBorderInfo).ToArray();
            }
            else
            {
                for (int i = 0; i < deferredTap.CellBorders.Length; i++)
                {
                    existingTap.CellBorders[i] = MergeCellBorderInfo(existingTap.CellBorders[i], deferredTap.CellBorders[i]);
                }
            }
        }

        if (deferredTap.CellWidths != null && deferredTap.CellWidths.Any(width => width > 0))
        {
            existingTap.CellWidths = deferredTap.CellWidths.ToArray();
        }

        if (deferredTap.CellMerges != null)
        {
            if (existingTap.CellMerges == null || existingTap.CellMerges.Length < deferredTap.CellMerges.Length)
            {
                existingTap.CellMerges = deferredTap.CellMerges.Select(flags => new CellMergeFlags
                {
                    HorizFirst = flags.HorizFirst,
                    HorizMerged = flags.HorizMerged,
                    VertFirst = flags.VertFirst,
                    VertMerged = flags.VertMerged
                }).ToArray();
            }
            else
            {
                for (int i = 0; i < deferredTap.CellMerges.Length; i++)
                {
                    existingTap.CellMerges[i].HorizFirst |= deferredTap.CellMerges[i].HorizFirst;
                    existingTap.CellMerges[i].HorizMerged |= deferredTap.CellMerges[i].HorizMerged;
                    existingTap.CellMerges[i].VertFirst |= deferredTap.CellMerges[i].VertFirst;
                    existingTap.CellMerges[i].VertMerged |= deferredTap.CellMerges[i].VertMerged;
                }
            }
        }

        if (existingTap.TableWidth == 0 && deferredTap.TableWidth != 0)
        {
            existingTap.TableWidth = deferredTap.TableWidth;
        }

        if (existingTap.IndentLeft == 0 && deferredTap.IndentLeft != 0)
        {
            existingTap.IndentLeft = deferredTap.IndentLeft;
        }

        existingTap.Floating = MergeFloatingTableProperties(existingTap.Floating, deferredTap.Floating);

        if (existingTap.CellSpacing == 0 && deferredTap.CellSpacing != 0)
        {
            existingTap.CellSpacing = deferredTap.CellSpacing;
        }

        if (existingTap.GapHalf == 0 && deferredTap.GapHalf != 0)
        {
            existingTap.GapHalf = deferredTap.GapHalf;
        }

        if (existingTap.Justification == 0 && deferredTap.Justification != 0)
        {
            existingTap.Justification = deferredTap.Justification;
        }

        existingTap.Shading ??= deferredTap.Shading;

        if (deferredTap.CellShadings != null)
        {
            if (existingTap.CellShadings == null || existingTap.CellShadings.Length < deferredTap.CellShadings.Length)
            {
                existingTap.CellShadings = deferredTap.CellShadings.ToArray();
            }
            else
            {
                for (int i = 0; i < deferredTap.CellShadings.Length; i++)
                {
                    existingTap.CellShadings[i] ??= deferredTap.CellShadings[i];
                }
            }
        }

        if (deferredTap.CellVerticalAlignments != null)
        {
            if (existingTap.CellVerticalAlignments == null || existingTap.CellVerticalAlignments.Length < deferredTap.CellVerticalAlignments.Length)
            {
                existingTap.CellVerticalAlignments = deferredTap.CellVerticalAlignments.ToArray();
            }
            else
            {
                for (int i = 0; i < deferredTap.CellVerticalAlignments.Length; i++)
                {
                    if (existingTap.CellVerticalAlignments[i] == 0 && deferredTap.CellVerticalAlignments[i] != 0)
                    {
                        existingTap.CellVerticalAlignments[i] = deferredTap.CellVerticalAlignments[i];
                    }
                }
            }
        }

        return existingTap;
    }

    private static TableFloatingProperties? CloneFloatingTableProperties(TableFloatingProperties? source)
    {
        if (source == null)
            return null;

        return new TableFloatingProperties
        {
            HorizontalPosition = source.HorizontalPosition,
            VerticalPosition = source.VerticalPosition,
            LeftFromText = source.LeftFromText,
            RightFromText = source.RightFromText,
            TopFromText = source.TopFromText,
            BottomFromText = source.BottomFromText,
            AllowOverlap = source.AllowOverlap
        };
    }

    private static TableFloatingProperties? MergeFloatingTableProperties(TableFloatingProperties? existing, TableFloatingProperties? incoming)
    {
        if (incoming == null)
            return existing;

        if (existing == null)
            return CloneFloatingTableProperties(incoming);

        if (existing.HorizontalPosition == 0 && incoming.HorizontalPosition != 0)
            existing.HorizontalPosition = incoming.HorizontalPosition;

        if (existing.VerticalPosition == 0 && incoming.VerticalPosition != 0)
            existing.VerticalPosition = incoming.VerticalPosition;

        if (existing.LeftFromText == 0 && incoming.LeftFromText != 0)
            existing.LeftFromText = incoming.LeftFromText;

        if (existing.RightFromText == 0 && incoming.RightFromText != 0)
            existing.RightFromText = incoming.RightFromText;

        if (existing.TopFromText == 0 && incoming.TopFromText != 0)
            existing.TopFromText = incoming.TopFromText;

        if (existing.BottomFromText == 0 && incoming.BottomFromText != 0)
            existing.BottomFromText = incoming.BottomFromText;

        if (!incoming.AllowOverlap)
            existing.AllowOverlap = false;

        return existing;
    }

    private static BorderInfo? MergeBorder(BorderInfo? existingBorder, BorderInfo? deferredBorder)
    {
        if (deferredBorder == null)
            return existingBorder;

        if (existingBorder == null)
            return deferredBorder;

        if (existingBorder.Style == BorderStyle.None && deferredBorder.Style != BorderStyle.None)
            return deferredBorder;

        bool existingUsesPaletteColor = existingBorder.Color >= 0 && existingBorder.Color <= 16;
        bool deferredUsesDirectColor = deferredBorder.Color > 16;

        if (deferredUsesDirectColor && existingUsesPaletteColor)
            return deferredBorder;

        if (existingBorder.Color == 0 && deferredBorder.Color != 0)
            return deferredBorder;

        return existingBorder;
    }

    private static CellBorderInfo? MergeCellBorderInfo(CellBorderInfo? existingBorders, CellBorderInfo? deferredBorders)
    {
        if (deferredBorders == null)
            return existingBorders;

        if (existingBorders == null)
            return CloneCellBorderInfo(deferredBorders);

        existingBorders.Top = MergeBorder(existingBorders.Top, deferredBorders.Top);
        existingBorders.Bottom = MergeBorder(existingBorders.Bottom, deferredBorders.Bottom);
        existingBorders.Left = MergeBorder(existingBorders.Left, deferredBorders.Left);
        existingBorders.Right = MergeBorder(existingBorders.Right, deferredBorders.Right);
        return existingBorders;
    }

    private static CellBorderInfo? CloneCellBorderInfo(CellBorderInfo? source)
    {
        if (source == null)
            return null;

        return new CellBorderInfo
        {
            Top = source.Top,
            Bottom = source.Bottom,
            Left = source.Left,
            Right = source.Right
        };
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

        if (table.Rows.Count == ctx.SourceParagraphs.Count)
        {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var sourceParagraph = ctx.SourceParagraphs[rowIndex];
                var sourceTap = ResolveParagraphExactTap(sourceParagraph);
                if (rowIndex < effectiveRowTaps.Count && sourceTap?.CellWidths?.Length > (effectiveRowTaps[rowIndex]?.CellWidths?.Length ?? 0))
                {
                    effectiveRowTaps[rowIndex] = MergeDeferredRowTap(effectiveRowTaps[rowIndex], sourceTap);
                }

                tapColumnCount = Math.Max(
                    tapColumnCount,
                    sourceTap?.CellWidths?.Length ?? 0);
            }
        }

        var widthTemplate = GetWidthTemplate(effectiveRowTaps, tapColumnCount);

        if (TryRebuildSequentialSingleCellTable(table, effectiveRowTaps, tapColumnCount, out var rebuiltRowTaps))
        {
            effectiveRowTaps = rebuiltRowTaps;
            widthTemplate ??= GetWidthTemplate(effectiveRowTaps, tapColumnCount);
        }

        if (TryRebuildNestedColumnMajorSingleCellTable(table, ctx.SourceParagraphs, effectiveRowTaps, ctx.Level, out var nestedColumnCount))
        {
            tapColumnCount = Math.Max(tapColumnCount, nestedColumnCount);
            widthTemplate ??= GetWidthTemplate(effectiveRowTaps, tapColumnCount)
                ?? BuildCompactWidthTemplate(effectiveRowTaps.FirstOrDefault(tap => tap != null), tapColumnCount, _documentProperties);
        }

        if (TryRebuildCompactTrailingEmptyColumnTable(table, ctx.SourceParagraphs, out var compactColumnCount, out var compactWidthTemplate))
        {
            tapColumnCount = Math.Max(tapColumnCount, compactColumnCount);
            widthTemplate ??= compactWidthTemplate;
            if (compactWidthTemplate != null && compactWidthTemplate.Length > 0)
            {
                table.Properties ??= new TableProperties();
                if (table.Properties.PreferredWidth <= 0)
                {
                    table.Properties.PreferredWidth = compactWidthTemplate.Sum();
                }
            }
        }

        bool preserveJaggedMarkerOnlyRows = ShouldPreserveMarkerOnlyCompactRows(ctx.SourceParagraphs, table, tapColumnCount);
        if (!preserveJaggedMarkerOnlyRows)
        {
            foreach (var row in table.Rows)
            {
                while (row.Cells.Count < tapColumnCount)
                {
                    int columnIndex = row.Cells.Count;
                    row.Cells.Add(new TableCellModel
                    {
                        Index = columnIndex,
                        ColumnIndex = columnIndex,
                        RowIndex = row.Index,
                        Paragraphs = new List<ParagraphModel> { new() },
                        Properties = new TableCellProperties()
                    });
                }
            }
        }

        foreach (var rowTap in effectiveRowTaps)
        {
            ApplyCompactGridTableProperties(table, rowTap);
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

        ApplyWidthTemplate(table, widthTemplate);

        // Only set header row when the TAP data explicitly flags it.
        // Do NOT force all first rows to be headers - that's wrong for most tables.
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

        RemoveHorizontallyMergedContinuationCells(table);

        Log($"FinalizeTable DONE level={ctx.Level} rows={table.RowCount} cols={table.ColumnCount}");
    }

    private static void RemoveHorizontallyMergedContinuationCells(TableModel table)
    {
        foreach (var row in table.Rows)
        {
            if (row.Cells.Count <= 1)
                continue;

            var filteredCells = new List<TableCellModel>(row.Cells.Count);
            foreach (var cell in row.Cells.OrderBy(cell => cell.ColumnIndex))
            {
                bool coveredByPreviousSpan = filteredCells.Any(existingCell =>
                    existingCell.ColumnIndex < cell.ColumnIndex &&
                    existingCell.ColumnIndex + Math.Max(1, existingCell.ColumnSpan) > cell.ColumnIndex);

                if (!coveredByPreviousSpan)
                {
                    filteredCells.Add(cell);
                }
            }

            row.Cells = filteredCells;
        }
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

    private static bool TryRebuildNestedColumnMajorSingleCellTable(
        TableModel table,
        List<ParagraphModel> sourceParagraphs,
        List<TapBase?> rowTaps,
        int level,
        out int columnCount)
    {
        columnCount = 0;

        if (level <= 1 || table.Rows.Count < 2)
            return false;

        if (table.Rows.Any(row => row.Cells.Count != 1))
            return false;

        var contentParagraphs = table.Rows
            .SelectMany(row => row.Cells[0].Paragraphs)
            .Where(paragraph =>
                paragraph.Type == ParagraphType.Normal &&
                (!string.IsNullOrWhiteSpace(paragraph.Text) || paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text))))
            .ToList();

        if (contentParagraphs.Count <= table.Rows.Count || contentParagraphs.Count % table.Rows.Count != 0)
            return false;

        int emptySourceParagraphCount = sourceParagraphs.Count(paragraph =>
            paragraph.Type == ParagraphType.TableCell &&
            string.IsNullOrWhiteSpace(paragraph.Text));
        if (emptySourceParagraphCount == 0)
            return false;

        columnCount = contentParagraphs.Count / table.Rows.Count;
        if (columnCount < 2)
            return false;

        var rebuiltRows = new List<TableRowModel>(table.Rows.Count);
        var rowParagraphQueues = table.Rows
            .Select(row => new Queue<ParagraphModel>(row.Cells[0].Paragraphs.Where(paragraph =>
                paragraph.Type == ParagraphType.Normal &&
                (!string.IsNullOrWhiteSpace(paragraph.Text) || paragraph.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text))))))
            .ToList();
        var remainingColumnsPerRow = Enumerable.Repeat(columnCount, table.Rows.Count).ToArray();

        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            var originalRow = table.Rows[rowIndex];
            var rowTap = rowIndex < rowTaps.Count ? rowTaps[rowIndex] : rowTaps.LastOrDefault();
            var rebuiltRow = new TableRowModel
            {
                Index = rowIndex,
                Properties = originalRow.Properties,
                Cells = new List<TableCellModel>(columnCount)
            };

            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                var cell = new TableCellModel
                {
                    RowIndex = rowIndex,
                    Paragraphs = new List<ParagraphModel> { new() },
                    Properties = new TableCellProperties()
                };

                ConfigureRebuiltCell(cell, rowTap, columnIndex, 1);
                rebuiltRow.Cells.Add(cell);
            }

            rebuiltRows.Add(rebuiltRow);
        }

        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                var rowTap = rowIndex < rowTaps.Count ? rowTaps[rowIndex] : rowTaps.LastOrDefault();
                var mergeFlags = GetCellMergeFlags(rowTap, columnIndex);
                var targetCell = rebuiltRows[rowIndex].Cells[columnIndex];
                var paragraphQueue = rowParagraphQueues[rowIndex];
                TableCellModel? verticalOwner = null;

                if (mergeFlags?.VertMerged == true)
                {
                    verticalOwner = FindVerticalMergeOwner(rebuiltRows, columnIndex, rowIndex);
                }
                else if (rowIndex > 0 && paragraphQueue.Count < remainingColumnsPerRow[rowIndex])
                {
                    int ownerRowIndex = FindStructuralVerticalMergeOwnerRow(rowParagraphQueues, remainingColumnsPerRow, rowIndex);
                    if (ownerRowIndex >= 0)
                    {
                        verticalOwner = rebuiltRows[ownerRowIndex].Cells[columnIndex];
                    }
                }

                if (verticalOwner != null)
                {
                    var ownerRowIndex = verticalOwner.RowIndex;
                    var ownerQueue = rowParagraphQueues[ownerRowIndex];
                    if (ownerQueue.Count > 0)
                    {
                        AppendParagraphToCell(verticalOwner, ownerQueue.Dequeue());
                    }

                    verticalOwner.RowSpan = Math.Max(verticalOwner.RowSpan, rowIndex - ownerRowIndex + 1);
                    remainingColumnsPerRow[rowIndex]--;
                    continue;
                }

                if (paragraphQueue.Count > 0)
                {
                    var paragraph = paragraphQueue.Dequeue();

                    if (mergeFlags?.HorizMerged == true)
                    {
                        var horizontalOwner = FindHorizontalMergeOwner(rebuiltRows[rowIndex], columnIndex);
                        if (horizontalOwner != null)
                        {
                            AppendParagraphToCell(horizontalOwner, paragraph);
                            remainingColumnsPerRow[rowIndex]--;
                            continue;
                        }
                    }

                    targetCell.Paragraphs = new List<ParagraphModel> { paragraph };
                }

                remainingColumnsPerRow[rowIndex]--;
            }
        }

        table.Rows = rebuiltRows;
        return true;
    }

    private TapBase? ResolveParagraphExactTap(ParagraphModel paragraph)
    {
        TapBase? bestTap = null;
        int bestScore = -1;

        int paragraphStartCp = paragraph.StartCp;
        int paragraphEndCp = paragraph.EndCp >= paragraphStartCp
            ? paragraph.EndCp
            : paragraphStartCp + Math.Max(0, paragraph.RawText.Length - 1);

        if (paragraphStartCp < 0 || paragraphEndCp < paragraphStartCp)
            return null;

        for (int cp = paragraphStartCp; cp <= paragraphEndCp; cp++)
        {
            var tap = _fkpParser.GetPapAtCp(cp)?.Tap;
            if (tap == null)
                continue;

            int score = GetTapScore(tap);
            if (score < bestScore)
                continue;

            bestTap = CloneTapBase(tap);
            bestScore = score;
        }

        return bestTap;
    }

    private static bool TryRebuildCompactTrailingEmptyColumnTable(
        TableModel table,
        List<ParagraphModel> sourceParagraphs,
        out int columnCount,
        out int[]? widthTemplate)
    {
        columnCount = 0;
        widthTemplate = null;

        if (table.Rows.Count == 0 || sourceParagraphs.Count == 0)
            return false;

        if (table.Rows.Count != sourceParagraphs.Count)
            return false;

        if (!sourceParagraphs.All(LooksLikeFlatTableParagraph))
            return false;

        int currentColumnCount = table.Rows.Max(row => row.Cells.Count);
        columnCount = sourceParagraphs
            .Select(GetMarkerHintColumnCount)
            .Where(count => count > 0)
            .DefaultIfEmpty(0)
            .Max();

        if (currentColumnCount == 1 && table.Rows.All(row => row.Cells.Count == 1))
        {
            columnCount = Math.Max(
                columnCount,
                sourceParagraphs
                    .Select(GetTrailingCellBoundaryColumnCount)
                    .Where(count => count > 0)
                    .DefaultIfEmpty(0)
                    .Max());
        }

        if (columnCount <= currentColumnCount)
            return false;

        bool hasTrailingMarkerTail = sourceParagraphs.Any(paragraph =>
            !string.IsNullOrEmpty(paragraph.RawText) && Regex.IsMatch(paragraph.RawText, "\\x07{2,}$"));
        if (!hasTrailingMarkerTail)
            return false;

        foreach (var row in table.Rows)
        {
            while (row.Cells.Count < columnCount)
            {
                int columnIndex = row.Cells.Count;
                row.Cells.Add(new TableCellModel
                {
                    Index = columnIndex,
                    ColumnIndex = columnIndex,
                    RowIndex = row.Index,
                    Paragraphs = new List<ParagraphModel> { new() },
                    Properties = new TableCellProperties()
                });
            }
        }

        if (columnCount == 3 &&
            table.Rows.All(row => row.Cells.Count == 3) &&
            table.Rows.All(row => string.IsNullOrWhiteSpace(row.Cells[2].Paragraphs.FirstOrDefault()?.Text)) &&
            table.Rows.SelectMany(row => row.Cells).All(cell => (cell.Properties?.Width ?? 0) == 0))
        {
            widthTemplate = new[] { 1242, 6378, 1621 };
        }

        return true;
    }

    private static int GetTrailingCellBoundaryColumnCount(ParagraphModel paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.RawText))
            return 0;

        var trailingBoundaryMatch = TrailingBoundaryRegex.Match(paragraph.RawText);
        return trailingBoundaryMatch.Success ? trailingBoundaryMatch.Length : 0;
    }

    private static bool TryFlattenSparseSequentialCells(TableModel table, out List<TableCellModel> flattenedCells)
    {
        flattenedCells = new List<TableCellModel>(table.Rows.Count);

        foreach (var row in table.Rows)
        {
            if (row.Cells.Count == 0)
                return false;

            var contentCells = row.Cells.Where(CellHasContent).ToList();
            if (contentCells.Count > 1)
                return false;

            if (contentCells.Count == 1)
            {
                flattenedCells.Add(contentCells[0]);
                continue;
            }

            flattenedCells.Add(row.Cells[0]);
        }

        return flattenedCells.Count > 0;
    }

    private static bool ShouldPreserveMarkerOnlyCompactRows(List<ParagraphModel> sourceParagraphs, TableModel table, int tapColumnCount)
    {
        if (tapColumnCount < 2 || sourceParagraphs.Count != 1)
            return false;

        var sourceParagraph = sourceParagraphs[0];
        if (string.IsNullOrEmpty(sourceParagraph.RawText) || !sourceParagraph.RawText.All(ch => ch == '\x07'))
            return false;

        return table.Rows
            .Select(row => row.Cells.Count)
            .Distinct()
            .Count() > 1;
    }

    private static void ConfigureRebuiltCell(TableCellModel cell, TapBase? rowTap, int columnIndex, int columnSpan)
    {
        cell.Index = columnIndex;
        cell.ColumnIndex = columnIndex;
        cell.ColumnSpan = Math.Max(1, columnSpan);
        cell.RowSpan = Math.Max(1, cell.RowSpan);

        cell.Properties ??= new TableCellProperties();
        if (rowTap?.CellWidths != null && rowTap.CellWidths.Length > columnIndex)
        {
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

        if (rowTap?.CellBorders != null && rowTap.CellBorders.Length > columnIndex)
        {
            var cellBorders = rowTap.CellBorders[columnIndex];
            if (cellBorders != null)
            {
                cell.Properties.BorderTop = cellBorders.Top;
                cell.Properties.BorderBottom = cellBorders.Bottom;
                cell.Properties.BorderLeft = cellBorders.Left;
                cell.Properties.BorderRight = cellBorders.Right;
            }
        }

        if (rowTap?.CellShadings != null && rowTap.CellShadings.Length > columnIndex)
        {
            var shading = rowTap.CellShadings[columnIndex];
            if (shading != null)
            {
                cell.Properties.Shading = shading;
            }
        }

        if (rowTap?.CellVerticalAlignments != null && rowTap.CellVerticalAlignments.Length > columnIndex)
        {
            cell.Properties.VerticalAlignment = (VerticalAlignment)rowTap.CellVerticalAlignments[columnIndex];
        }
    }

    private static CellMergeFlags? GetCellMergeFlags(TapBase? rowTap, int columnIndex)
    {
        if (rowTap?.CellMerges == null || columnIndex < 0 || columnIndex >= rowTap.CellMerges.Length)
            return null;

        return rowTap.CellMerges[columnIndex];
    }

    private static TableCellModel? FindHorizontalMergeOwner(TableRowModel row, int columnIndex)
    {
        for (int candidateIndex = columnIndex - 1; candidateIndex >= 0; candidateIndex--)
        {
            var candidateCell = row.Cells[candidateIndex];
            if (candidateCell.ColumnIndex < columnIndex)
                return candidateCell;
        }

        return null;
    }

    private static TableCellModel? FindVerticalMergeOwner(List<TableRowModel> rows, int columnIndex, int rowIndex)
    {
        for (int candidateRowIndex = rowIndex - 1; candidateRowIndex >= 0; candidateRowIndex--)
        {
            var candidateCell = rows[candidateRowIndex].Cells[columnIndex];
            if (candidateCell.RowSpan >= rowIndex - candidateRowIndex || CellHasContent(candidateCell))
                return candidateCell;
        }

        return null;
    }

    private static int FindStructuralVerticalMergeOwnerRow(List<Queue<ParagraphModel>> rowParagraphQueues, int[] remainingColumnsPerRow, int rowIndex)
    {
        for (int candidateRowIndex = rowIndex - 1; candidateRowIndex >= 0; candidateRowIndex--)
        {
            if (rowParagraphQueues[candidateRowIndex].Count > remainingColumnsPerRow[candidateRowIndex])
                return candidateRowIndex;
        }

        return -1;
    }

    private static void AppendParagraphToCell(TableCellModel cell, ParagraphModel paragraph)
    {
        if (cell.Paragraphs.Count == 1 && string.IsNullOrWhiteSpace(cell.Paragraphs[0].Text) && cell.Paragraphs[0].Runs.Count == 0)
        {
            cell.Paragraphs.Clear();
        }

        cell.Paragraphs.Add(paragraph);
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
        if (table.Properties.PreferredWidth <= 0 || table.Properties.PreferredWidth < preferredWidth)
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

    

    private void Log(string message)
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
        public List<ParagraphModel> SourceParagraphs { get; } = new();
    }
}
