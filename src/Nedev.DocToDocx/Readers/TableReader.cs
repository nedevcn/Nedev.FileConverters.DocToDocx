using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Readers;

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
                            Text = r.Text?.Replace("\x07", ""),
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
            // Empty cell? 
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
