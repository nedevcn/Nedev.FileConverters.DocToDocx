using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Readers;

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
        var topLevelTables = new List<TableModel>();
        var stack = new List<TableContext>(); // index 0 = level 1, index 1 = level 2, etc.

        foreach (var para in document.Paragraphs.OrderBy(p => p.Index))
        {
            int level = para.NestingLevel > 0 ? para.NestingLevel : 0;
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
                
                FlushCurrentCell(popped);
                FlushCurrentRow(popped);
                FinalizeTable(popped);

                // If the popped table actually has content, attach it
                if (popped.Table.Rows.Count > 0)
                {
                    if (stack.Count > 0)
                    {
                        var parent = stack.Last();
                        // record parent index for easier navigation later
                        popped.Table.ParentTableIndex = parent.Table.Index;
                        
                        // Attach to current cell of parent
                        var nestedPara = new ParagraphModel
                        {
                            Type = ParagraphType.NestedTable,
                            NestedTable = popped.Table,
                            NestingLevel = parent.Level
                        };
                        parent.CurrentCellParagraphs.Add(nestedPara);
                    }
                    else
                    {
                        // top-level table has no parent
                        popped.Table.ParentTableIndex = null;
                        topLevelTables.Add(popped.Table);
                    }
                }
            }

            if (level == 0) continue;

            // Ensure we have active contexts up to the current level
            while (stack.Count < level)
            {
                var newCtx = new TableContext
                {
                    Level = stack.Count + 1,
                    Table = new TableModel { Index = topLevelTables.Count, StartParagraphIndex = para.Index }
                };
                stack.Add(newCtx);
            }

            var activeCtx = stack.Last();
            activeCtx.LastTableParagraphIndex = para.Index;

            // Extract TAP properties for table alignment and cell widths
            TapBase? tapForParagraph = null;
            var firstRun = para.Runs.FirstOrDefault();
            if (firstRun != null)
            {
                var pap = _fkpParser.GetPapAtCp(firstRun.CharacterPosition);
                tapForParagraph = pap?.Tap;

                if (tapForParagraph != null && activeCtx.Table.Properties == null)
                {
                    activeCtx.Table.Properties = new TableProperties
                    {
                        Alignment = tapForParagraph.Justification switch
                        {
                            1 => TableAlignment.Center,
                            2 => TableAlignment.Right,
                            _ => TableAlignment.Left
                        },
                        CellSpacing = tapForParagraph.CellSpacing != 0
                            ? tapForParagraph.CellSpacing
                            : (tapForParagraph.GapHalf != 0 ? tapForParagraph.GapHalf * 2 : 0),
                        Indent = tapForParagraph.IndentLeft,
                        PreferredWidth = tapForParagraph.TableWidth,
                        BorderTop = tapForParagraph.BorderTop,
                        BorderBottom = tapForParagraph.BorderBottom,
                        BorderLeft = tapForParagraph.BorderLeft,
                        BorderRight = tapForParagraph.BorderRight,
                        BorderInsideH = tapForParagraph.BorderInsideH,
                        BorderInsideV = tapForParagraph.BorderInsideV,
                        Shading = tapForParagraph.Shading
                    };
                }

                if (activeCtx.CurrentRowTap == null && tapForParagraph != null)
                {
                    activeCtx.CurrentRowTap = tapForParagraph;
                }
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
            if (popped.Table.Rows.Count > 0)
            {
                if (stack.Count > 0)
                {
                    var parent = stack.Last();
                    popped.Table.ParentTableIndex = parent.Table.Index;
                    parent.CurrentCellParagraphs.Add(new ParagraphModel
                    {
                        Type = ParagraphType.NestedTable,
                        NestedTable = popped.Table,
                        NestingLevel = parent.Level
                    });
                }
                else
                {
                    popped.Table.ParentTableIndex = null;
                    topLevelTables.Add(popped.Table);
                }
            }
        }

        document.Tables = topLevelTables;

        if (NeedsFlatTableRecovery(document))
        {
            RecoverFlatTables(document);
        }
    }

    private static bool NeedsFlatTableRecovery(DocumentModel document)
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
            return;

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

        if (document.Tables.Count > 0 && document.Paragraphs.Count > 0 && document.Tables[0].StartParagraphIndex == 1)
        {
            var title = document.Paragraphs[0];
            if (!string.IsNullOrWhiteSpace(title.Text) && title.Text.Length <= 40)
            {
                title.Properties ??= new ParagraphProperties();
                title.Properties.Alignment = ParagraphAlignment.Center;
                title.Properties.KeepWithNext = true;
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
            table.Rows.Add(BuildRecoveredRow(rowCells, table.Rows.Count, table.ColumnCount, paragraphsToRemove));
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

    private static int EstimateColumnCount(IEnumerable<ParagraphModel> paragraphs)
    {
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

    private static (TableModel? Table, List<ParagraphModel> TrailingParagraphs) BuildFlatTable(
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
        var pendingCells = new List<string>();

        for (int paragraphIndex = 0; paragraphIndex < group.Count; paragraphIndex++)
        {
            var paragraph = group[paragraphIndex];
            var rowCandidates = GetRowCandidates(paragraph).ToList();
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

                foreach (var cell in cells)
                {
                    pendingCells.Add(cell);
                    if (pendingCells.Count == columnCount)
                    {
                        table.Rows.Add(BuildRecoveredRow(pendingCells, rowIndex++, columnCount, group));
                        pendingCells.Clear();
                    }
                }
            }
        }

        if (pendingCells.Count == columnCount)
        {
            table.Rows.Add(BuildRecoveredRow(pendingCells, rowIndex, columnCount, group));
            pendingCells.Clear();
        }

        if (pendingCells.Count > 0)
        {
            trailingParagraphs.AddRange(BuildTrailingParagraphs(string.Join("\r", pendingCells), group.LastOrDefault()));
        }

        table.RowCount = table.Rows.Count;
        if (table.RowCount == 0)
            return (null, new List<ParagraphModel>());

        return (table, trailingParagraphs);
    }

    private static TableRowModel BuildRecoveredRow(List<string> cellTexts, int rowIndex, int columnCount, List<ParagraphModel> sourceParagraphs)
    {
        var row = new TableRowModel { Index = rowIndex };
        int cellWidth = 9360 / Math.Max(1, columnCount);
        var sourceProps = sourceParagraphs.FirstOrDefault()?.Properties;
        var sourceRunProps = sourceParagraphs.FirstOrDefault(p => p.Runs.Count > 0)?.Runs[0].Properties;

        for (int columnIndex = 0; columnIndex < columnCount && columnIndex < cellTexts.Count; columnIndex++)
        {
            string cellText = cellTexts[columnIndex].Trim();
            var paragraph = new ParagraphModel
            {
                Type = ParagraphType.Normal,
                Properties = sourceProps,
                Runs = new List<RunModel>()
            };

            if (!string.IsNullOrEmpty(cellText))
            {
                paragraph.Runs.Add(new RunModel
                {
                    Text = cellText,
                    Properties = sourceRunProps
                });
            }

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

    private static IEnumerable<List<string>> GetRowCandidates(ParagraphModel paragraph)
    {
        if (string.IsNullOrEmpty(paragraph.RawText))
            yield break;

        foreach (var rawLine in paragraph.RawText.Split('\r', StringSplitOptions.RemoveEmptyEntries))
        {
            var cells = rawLine
                .Split('\x07')
                .Select(NormalizeFlatCellText)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            if (cells.Count > 0)
                yield return cells;
        }
    }

    private static string NormalizeFlatCellText(string text)
    {
        return text
            .Replace("\r", "\n")
            .Trim('\r', '\n', '\t', ' ');
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

    private static string SplitTrailingText(ref List<string> cells, int columnCount)
    {
        if (cells.Count == 0)
            return string.Empty;

        if (cells.Count > columnCount)
        {
            var trailingOverflow = cells.Skip(columnCount).ToList();
            cells = cells.Take(columnCount).ToList();
            return string.Join("\r", trailingOverflow);
        }

        if (cells.Count < columnCount)
            return string.Empty;

        var lastCell = cells[columnCount - 1];
        var match = Regex.Match(lastCell, "(?<cell>.*?)(?<trail>(第[一二三四五六七八九十百0-9]+条[：:].*|[0-9]+\\.[0-9].*))$");
        if (!match.Success)
            return string.Empty;

        var cellText = match.Groups["cell"].Value.Trim();
        var trailingText = match.Groups["trail"].Value.Trim();
        if (string.IsNullOrWhiteSpace(cellText) || string.IsNullOrWhiteSpace(trailingText))
            return string.Empty;

        cells[columnCount - 1] = cellText;
        return trailingText;
    }

    private static List<ParagraphModel> BuildTrailingParagraphs(string trailingText, ParagraphModel? sourceParagraph)
    {
        var paragraphs = new List<ParagraphModel>();
        if (string.IsNullOrWhiteSpace(trailingText))
            return paragraphs;

        foreach (var part in trailingText.Split('\r', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (string.IsNullOrWhiteSpace(part))
                continue;

            var paragraph = new ParagraphModel
            {
                Type = ParagraphType.Normal,
                RawText = part,
                Properties = sourceParagraph?.Properties,
                Runs = new List<RunModel>()
            };
            paragraph.Runs.Add(new RunModel
            {
                Text = part,
                Properties = sourceParagraph?.Runs.FirstOrDefault()?.Properties
            });
            paragraphs.Add(paragraph);
        }

        return paragraphs;
    }

    private void FlushCurrentCell(TableContext ctx)
    {
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

        if (ctx.CurrentRowTap?.CellWidths != null && ctx.CurrentRowTap.CellWidths.Length > cellModel.ColumnIndex)
        {
            cellModel.Properties ??= new TableCellProperties();
            cellModel.Properties.Width = ctx.CurrentRowTap.CellWidths[cellModel.ColumnIndex];
        }

        ctx.CellsInCurrentRow.Add(cellModel);
        ctx.CurrentCellParagraphs.Clear();
    }

    private void FlushCurrentRow(TableContext ctx)
    {
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
    }

    private void FinalizeTable(TableContext ctx)
    {
        var table = ctx.Table;
        if (table.Rows.Count == 0) return;

        table.EndParagraphIndex = ctx.LastTableParagraphIndex;
        table.RowCount = table.Rows.Count;
        table.ColumnCount = table.Rows.Max(r => r.Cells.Count);

        // Only set header row when the TAP data explicitly flags it.
        // Do NOT force all first rows to be headers — that's wrong for most tables.
        var firstRow = table.Rows.FirstOrDefault();
        var firstTap = ctx.RowTaps.Count > 0 ? ctx.RowTaps[0] : null;
        if (firstRow != null && firstTap != null && firstTap.IsHeaderRow)
        {
            firstRow.Properties ??= new TableRowProperties();
            firstRow.Properties.IsHeaderRow = true;
        }

        // Apply Spans
        bool hasTapMergeInfo = ctx.RowTaps.Any(t => t?.CellMerges != null);
        if (hasTapMergeInfo && table.ColumnCount > 0)
        {
            for (int col = 0; col < table.ColumnCount; col++)
            {
                int row = 0;
                while (row < table.Rows.Count)
                {
                    var startCell = GetCell(table, row, col);
                    if (startCell == null) { row++; continue; }

                    var tap = row < ctx.RowTaps.Count ? ctx.RowTaps[row] : null;
                    var flags = tap?.CellMerges != null && col < tap.CellMerges.Length ? tap.CellMerges[col] : null;
                    if (flags == null || !flags.VertFirst) { row++; continue; }

                    int span = 1;
                    int nextRow = row + 1;
                    while (nextRow < table.Rows.Count)
                    {
                        var nextTap = nextRow < ctx.RowTaps.Count ? ctx.RowTaps[nextRow] : null;
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
                var tap = row < ctx.RowTaps.Count ? ctx.RowTaps[row] : null;
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
        else if (table.ColumnCount > 0)
        {
            // Heuristic vertically merges empty cells below a content cell
            for (int col = 0; col < table.ColumnCount; col++)
            {
                int row = 0;
                while (row < table.Rows.Count)
                {
                    var startCell = GetCell(table, row, col);
                    if (startCell == null || !CellHasContent(startCell)) { row++; continue; }

                    int span = 1;
                    int nextRow = row + 1;
                    while (nextRow < table.Rows.Count)
                    {
                        var nextCell = GetCell(table, nextRow, col);
                        if (nextCell == null || CellHasContent(nextCell)) break;
                        span++;
                        nextRow++;
                    }

                    if (span > 1) { startCell.RowSpan = span; row += span; }
                    else { row++; }
                }
            }
        }
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
