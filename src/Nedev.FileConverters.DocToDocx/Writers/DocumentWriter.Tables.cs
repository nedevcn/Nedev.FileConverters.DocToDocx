using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// DocumentWriter partial class — table writing methods.
/// </summary>
public partial class DocumentWriter
{
    /// <summary>
    /// Writes a table to the document.
    /// </summary>
    private void WriteTable(TableModel table)
    {
        Logger.Debug($"DocumentWriter.WriteTable START: startPara={table.StartParagraphIndex} endPara={table.EndParagraphIndex} columns={table.ColumnCount} rows={table.Rows.Count} IsNested={table.IsNested} ParentTableIndex={table.ParentTableIndex}");
        
        // Log first cell content for debugging
        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count > 0)
        {
            var firstCell = table.Rows[0].Cells[0];
            var firstCellText = string.Join("; ", firstCell.Paragraphs.Select(p => p.Text));
            var nestedTableCount = firstCell.Paragraphs.Count(p => p.Type == ParagraphType.NestedTable && p.NestedTable != null);
            Logger.Debug($"DocumentWriter.WriteTable: First cell has {firstCell.Paragraphs.Count} paragraphs, {nestedTableCount} nested tables. Text = '{firstCellText}'");
        }
        
        _writer.WriteStartElement("w", "tbl", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Write table properties
        WriteTableProperties(table);
        
        _writer.WriteStartElement("w", "tblGrid", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        int columnCount = table.ColumnCount > 0
            ? table.ColumnCount
            : (table.Rows.Any() ? table.Rows.Max(r => r.Cells.Count) : 0);
        
        var columnWidths = CalculateColumnWidths(table, columnCount);
        
        for (int i = 0; i < columnCount; i++)
        {
            _writer.WriteStartElement("w", "gridCol", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            int width = columnWidths[i];
            
            if (width > 0)
            {
                _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", width.ToString());
            }
            _writer.WriteEndElement();
        }
        _writer.WriteEndElement();

        // Write each row
        foreach (var row in table.Rows)
        {
            WriteTableRow(row, table, columnWidths);
        }
        
        _writer.WriteEndElement(); // w:tbl
    }

    private int[] CalculateColumnWidths(TableModel table, int columnCount)
    {
        var columnWidths = new int[columnCount];

        foreach (var row in table.Rows)
        {
            foreach (var cell in row.Cells)
            {
                if (cell.Properties?.Width is not > 0 || cell.ColumnIndex < 0 || cell.ColumnIndex >= columnCount)
                    continue;

                int span = Math.Max(1, cell.ColumnSpan);
                if (span == 1)
                {
                    columnWidths[cell.ColumnIndex] = Math.Max(columnWidths[cell.ColumnIndex], cell.Properties.Width);
                }
            }
        }

        foreach (var row in table.Rows)
        {
            foreach (var cell in row.Cells)
            {
                if (cell.Properties?.Width is not > 0 || cell.ColumnIndex < 0 || cell.ColumnIndex >= columnCount)
                    continue;

                int span = Math.Max(1, cell.ColumnSpan);
                if (span == 1)
                    continue;

                int endColumn = Math.Min(columnCount, cell.ColumnIndex + span);
                int knownWidth = 0;
                int unknownCount = 0;
                for (int columnIndex = cell.ColumnIndex; columnIndex < endColumn; columnIndex++)
                {
                    if (columnWidths[columnIndex] > 0)
                    {
                        knownWidth += columnWidths[columnIndex];
                    }
                    else
                    {
                        unknownCount++;
                    }
                }

                if (unknownCount == 0)
                    continue;

                int remainingWidth = cell.Properties.Width - knownWidth;
                if (remainingWidth <= 0)
                    continue;

                int widthPerUnknownColumn = Math.Max(1, remainingWidth / unknownCount);
                int remainder = Math.Max(0, remainingWidth % unknownCount);
                for (int columnIndex = cell.ColumnIndex; columnIndex < endColumn; columnIndex++)
                {
                    if (columnWidths[columnIndex] > 0)
                        continue;

                    columnWidths[columnIndex] = widthPerUnknownColumn + (remainder > 0 ? 1 : 0);
                    if (remainder > 0)
                    {
                        remainder--;
                    }
                }
            }
        }

        if (columnWidths.All(width => width == 0) && TryInferCalendarColumnWidths(table, out var inferredColumnWidths))
        {
            return inferredColumnWidths;
        }

        return columnWidths;
    }
        
    private bool TryInferCalendarColumnWidths(TableModel table, out int[] columnWidths)
    {
        columnWidths = Array.Empty<int>();

        if (table.ColumnCount != 13 || table.Rows.Count < 2 || table.Rows[0].Cells.Count != 1 || table.Rows[1].Cells.Count != 13)
            return false;

        var title = table.Rows[0].Cells[0].Paragraphs.FirstOrDefault()?.Text;
        if (string.IsNullOrWhiteSpace(title) || !System.Text.RegularExpressions.Regex.IsMatch(title.Trim(), "^(January|February|March|April|May|June|July|August|September|October|November|December)\\s+\\d{4}$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
            return false;

        string[] expectedDays = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
        var headerTexts = table.Rows[1].Cells.Select(cell => cell.Paragraphs.FirstOrDefault()?.Text ?? string.Empty).ToList();
        if (!headerTexts.Where((_, index) => index % 2 == 0).SequenceEqual(expectedDays, StringComparer.OrdinalIgnoreCase))
            return false;

        if (headerTexts.Where((_, index) => index % 2 == 1).Any(text => !string.IsNullOrWhiteSpace(text)))
            return false;

        int pageWidth = _document?.Properties.PageWidth ?? 12240;
        int marginLeft = _document?.Properties.MarginLeft ?? 1440;
        int marginRight = _document?.Properties.MarginRight ?? 1440;
        int preferredWidth = table.Properties?.PreferredWidth > 0
            ? table.Properties.PreferredWidth
            : Math.Max(1, pageWidth - marginLeft - marginRight);
        preferredWidth = Math.Max(1, preferredWidth);

        int separatorCount = 6;
        int contentCount = 7;
        int separatorWidth = table.Properties?.CellSpacing ?? 0;
        if (separatorWidth <= 0)
        {
            separatorWidth = Math.Max(1, preferredWidth / Math.Max(1, (contentCount * 4) + separatorCount));
        }

        int contentWidth = Math.Max(1, (preferredWidth - (separatorCount * separatorWidth)) / contentCount);
        int remaining = preferredWidth - (contentWidth * contentCount) - (separatorWidth * separatorCount);

        columnWidths = new int[13];
        for (int index = 0; index < columnWidths.Length; index++)
        {
            columnWidths[index] = index % 2 == 0 ? contentWidth : separatorWidth;
        }

        for (int index = 1; index < columnWidths.Length && remaining > 0; index += 2, remaining--)
        {
            columnWidths[index]++;
        }

        for (int index = 0; index < columnWidths.Length && remaining > 0; index += 2, remaining--)
        {
            columnWidths[index]++;
        }

        return columnWidths.All(width => width > 0);
    }

    /// <summary>
    /// Writes table properties (tblPr).
    /// </summary>
    private void WriteTableProperties(TableModel table)
    {
        var tableProperties = ResolveEffectiveTableProperties(table);

        _writer.WriteStartElement("w", "tblPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Table style
        if (tableProperties?.StyleIndex >= 0)
        {
            var style = _document?.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Table && s.StyleId == tableProperties.StyleIndex);
            var styleId = StyleHelper.GetTableStyleId(tableProperties.StyleIndex, style?.Name);
            
            _writer.WriteStartElement("w", "tblStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", styleId);
            _writer.WriteEndElement();
        }
        
        // Table width: prefer an explicit width from TAP when available, otherwise
        // let Word auto-size based on content.
        _writer.WriteStartElement("w", "tblW", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var preferredWidth = tableProperties?.PreferredWidth ?? 0;
        if (preferredWidth > 0)
        {
            preferredWidth = Math.Clamp(preferredWidth, 1, 31680);
            _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", preferredWidth.ToString());
            _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        }
        else
        {
            _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0");
            _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "auto");
        }
        _writer.WriteEndElement();
        
        // Table justification (alignment)
        if (tableProperties != null && tableProperties.Alignment != TableAlignment.Left)
        {
            _writer.WriteStartElement("w", "jc", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var alignment = tableProperties.Alignment switch
            {
                TableAlignment.Center => "center",
                TableAlignment.Right => "right",
                _ => "left"
            };
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", alignment);
            _writer.WriteEndElement();
        }
        
        // Table indent from left margin, when specified. This mirrors sprmTDxaLeft
        // and helps nested or offset tables align closer to the original layout.
        if (tableProperties != null && tableProperties.Indent != 0)
        {
            var clampedIndent = Math.Clamp(tableProperties.Indent, -31680, 31680);
            _writer.WriteStartElement("w", "tblInd", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", clampedIndent.ToString());
            _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
            _writer.WriteEndElement();
        }
        
        // Table borders
        if (tableProperties?.BorderTop != null || tableProperties?.BorderBottom != null ||
            tableProperties?.BorderLeft != null || tableProperties?.BorderRight != null ||
            tableProperties?.BorderInsideH != null || tableProperties?.BorderInsideV != null)
        {
            _writer.WriteStartElement("w", "tblBorders", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (tableProperties.BorderTop != null) WriteBorder("top", tableProperties.BorderTop);
            if (tableProperties.BorderBottom != null) WriteBorder("bottom", tableProperties.BorderBottom);
            if (tableProperties.BorderLeft != null) WriteBorder("left", tableProperties.BorderLeft);
            if (tableProperties.BorderRight != null) WriteBorder("right", tableProperties.BorderRight);
            if (tableProperties.BorderInsideH != null) WriteBorder("insideH", tableProperties.BorderInsideH);
            if (tableProperties.BorderInsideV != null) WriteBorder("insideV", tableProperties.BorderInsideV);
            _writer.WriteEndElement();
        }
        
        // Table shading
        if (tableProperties?.Shading != null)
        {
            WriteShading(tableProperties.Shading);
        }
        
        // Table cell margin: when the TAP exposes an inter-cell spacing we map it
        // to symmetric left/right padding; otherwise we fall back to a sensible
        // default that keeps existing documents visually similar.
        var spacing = tableProperties?.CellSpacing ?? 0;
        // Clamp to a small, non-negative range so extreme values from corrupted
        // documents do not explode layout.
        if (spacing < 0) spacing = 0;
        if (spacing > 720) spacing = 720; // max 0.5"

        int sidePadding = spacing > 0 ? spacing / 2 : 108;

        _writer.WriteStartElement("w", "tblCellMar", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteStartElement("w", "top", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0");
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "left", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", sidePadding.ToString());
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "bottom", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "0");
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "right", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", sidePadding.ToString());
        _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement(); // w:tblPr
    }

    private TableProperties? ResolveEffectiveTableProperties(TableModel table)
    {
        TableProperties? resolved = null;

        if (table.Properties != null)
        {
            resolved = CloneTableProperties(table.Properties);
        }

        if (resolved?.StyleIndex >= 0)
        {
            var styleProps = _document?.Styles.Styles
                .FirstOrDefault(style => style.Type == StyleType.Table && style.StyleId == resolved.StyleIndex)?
                .TableProperties;

            if (styleProps != null)
            {
                resolved.MergeWith(styleProps);
            }
        }

        if (resolved == null && !TableNeedsVisibleBorders(table))
        {
            return null;
        }

        resolved ??= new TableProperties();

        if (!HasAnyTableBorders(resolved) && TableNeedsVisibleBorders(table))
        {
            var fallbackBorder = new BorderInfo
            {
                Style = BorderStyle.Single,
                Width = 4,
                Space = 0,
                Color = 0
            };

            resolved.BorderTop = fallbackBorder;
            resolved.BorderBottom = fallbackBorder;
            resolved.BorderLeft = fallbackBorder;
            resolved.BorderRight = fallbackBorder;
            resolved.BorderInsideH = fallbackBorder;
            resolved.BorderInsideV = fallbackBorder;
        }

        return resolved;
    }

    private static bool TableNeedsVisibleBorders(TableModel table)
    {
        foreach (var row in table.Rows)
        {
            foreach (var cell in row.Cells)
            {
                foreach (var paragraph in cell.Paragraphs)
                {
                    if (paragraph.Type == ParagraphType.NestedTable && paragraph.NestedTable != null)
                    {
                        return true;
                    }

                    if (paragraph.Runs.Any(HasRenderableContent))
                    {
                        return true;
                    }
                }
            }
        }

        return false;
    }

    private static bool HasAnyTableBorders(TableProperties props)
    {
        return props.BorderTop != null ||
               props.BorderBottom != null ||
               props.BorderLeft != null ||
               props.BorderRight != null ||
               props.BorderInsideH != null ||
               props.BorderInsideV != null;
    }

    private static TableProperties CloneTableProperties(TableProperties source)
    {
        return new TableProperties
        {
            StyleIndex = source.StyleIndex,
            CellSpacing = source.CellSpacing,
            Indent = source.Indent,
            Alignment = source.Alignment,
            PreferredWidth = source.PreferredWidth,
            BorderTop = source.BorderTop,
            BorderBottom = source.BorderBottom,
            BorderLeft = source.BorderLeft,
            BorderRight = source.BorderRight,
            BorderInsideH = source.BorderInsideH,
            BorderInsideV = source.BorderInsideV,
            Shading = source.Shading
        };
    }

    /// <summary>
    /// Writes a table row.
    /// </summary>
    private void WriteTableRow(TableRowModel row, TableModel table, int[] columnWidths)
    {
        Logger.Debug($"DocumentWriter.WriteTableRow START: tableStart={table.StartParagraphIndex} rowIndex={row.Index} cells={row.Cells.Count}");
        _writer.WriteStartElement("w", "tr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Row properties
        if (row.Properties != null)
        {
            _writer.WriteStartElement("w", "trPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            if (row.Properties.Height > 0)
            {
                var rowHeight = Math.Clamp(row.Properties.Height, 1, 31680);
                _writer.WriteStartElement("w", "trHeight", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", rowHeight.ToString());
                if (row.Properties.HeightIsExact)
                {
                    _writer.WriteAttributeString("w", "hRule", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "exact");
                }
                _writer.WriteEndElement();
            }
            
            if (row.Properties.IsHeaderRow)
            {
                _writer.WriteStartElement("w", "tblHeader", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteEndElement();
            }

            // Prevent row from being split across pages when requested
            if (!row.Properties.AllowBreakAcrossPages)
            {
                _writer.WriteStartElement("w", "cantSplit", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteEndElement();
            }
            
            _writer.WriteEndElement(); // w:trPr
        }
        
        // Write each cell
        foreach (var cell in row.Cells)
        {
            if (IsCoveredByHorizontalMerge(table, row.Index, cell.ColumnIndex))
            {
                // In OOXML, horizontally covered cells (due to gridSpan) MUST NOT be output as w:tc elements.
                // Otherwise, the number of columns in the row exceeds the table grid, causing severe corruption.
                continue;
            }
            WriteTableCell(cell, row, table, columnWidths);
        }
        
        _writer.WriteEndElement(); // w:tr
    }

    /// <summary>
    /// Writes a table cell, including vertical (vMerge) and horizontal (gridSpan)
    /// merges. For vertical merges we emit w:vMerge restart/continue based on
    /// RowSpan and cells in previous rows; for horizontal merges we emit
    /// w:gridSpan on the first cell and suppress content in covered cells.
    /// </summary>
    private void WriteTableCell(TableCellModel cell, TableRowModel row, TableModel table, int[] columnWidths)
    {
        _writer.WriteStartElement("w", "tc", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // Determine vertical merge role for this cell
        bool isVmergeStart = cell.RowSpan > 1;
        bool isVmergeContinue = !isVmergeStart && IsCoveredByVerticalMerge(table, row.Index, cell.ColumnIndex);
        
        int effectiveCellWidth = cell.Properties?.Width ?? 0;
        if (effectiveCellWidth <= 0)
        {
            int startColumn = Math.Clamp(cell.ColumnIndex, 0, Math.Max(0, columnWidths.Length - 1));
            int endColumn = Math.Min(columnWidths.Length, startColumn + Math.Max(1, cell.ColumnSpan));
            for (int columnIndex = startColumn; columnIndex < endColumn; columnIndex++)
            {
                effectiveCellWidth += columnWidths[columnIndex];
            }
        }

        bool hasTcPr = effectiveCellWidth > 0 || cell.ColumnSpan > 1 || cell.RowSpan > 1 || isVmergeContinue ||
                       cell.Properties?.BorderTop != null || cell.Properties?.BorderBottom != null ||
                       cell.Properties?.BorderLeft != null || cell.Properties?.BorderRight != null ||
                       cell.Properties?.Shading != null ||
                       cell.Properties?.NoWrap == true ||
                       (cell.Properties != null && cell.Properties.VerticalAlignment != VerticalAlignment.Top);

        if (hasTcPr)
        {
            // tcPr: tcW -> gridSpan -> vMerge -> tcBorders -> shd -> noWrap -> vAlign
            _writer.WriteStartElement("w", "tcPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            // Cell width
            if (effectiveCellWidth > 0)
            {
                var cellWidth = Math.Clamp(effectiveCellWidth, 1, 31680);
                _writer.WriteStartElement("w", "tcW", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", cellWidth.ToString());
                _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
                _writer.WriteEndElement();
            }
            
            // Grid span (column span) — only on the first (uncovered) cell
            if (cell.ColumnSpan > 1)
            {
                _writer.WriteStartElement("w", "gridSpan", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", cell.ColumnSpan.ToString());
                _writer.WriteEndElement();
            }
            
            // Vertical merge (row span)
            if (isVmergeStart || isVmergeContinue)
            {
                _writer.WriteStartElement("w", "vMerge", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                if (isVmergeStart)
                {
                    _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "restart");
                }
                _writer.WriteEndElement();
            }
            
            // Cell borders
            if (cell.Properties?.BorderTop != null || cell.Properties?.BorderBottom != null ||
                cell.Properties?.BorderLeft != null || cell.Properties?.BorderRight != null)
            {
                _writer.WriteStartElement("w", "tcBorders", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                if (cell.Properties.BorderTop != null) WriteBorder("top", cell.Properties.BorderTop);
                if (cell.Properties.BorderBottom != null) WriteBorder("bottom", cell.Properties.BorderBottom);
                if (cell.Properties.BorderLeft != null) WriteBorder("left", cell.Properties.BorderLeft);
                if (cell.Properties.BorderRight != null) WriteBorder("right", cell.Properties.BorderRight);
                _writer.WriteEndElement();
            }

            // Cell shading (shd)
            if (cell.Properties?.Shading != null)
            {
                WriteShading(cell.Properties.Shading);
            }

            // No wrap
            if (cell.Properties?.NoWrap == true)
            {
                _writer.WriteStartElement("w", "noWrap", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteEndElement();
            }

            // Vertical alignment
            if (cell.Properties != null && cell.Properties.VerticalAlignment != VerticalAlignment.Top)
            {
                _writer.WriteStartElement("w", "vAlign", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                var vAlign = cell.Properties.VerticalAlignment switch
                {
                    VerticalAlignment.Center => "center",
                    VerticalAlignment.Bottom => "bottom",
                    VerticalAlignment.Both => "both",
                    _ => "top"
                };
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", vAlign);
                _writer.WriteEndElement();
            }
            
            _writer.WriteEndElement(); // w:tcPr
        }
        
        // Write cell content (paragraphs) only for non-vMerge-continue cells.
        // Horizontal-merge covered cells are skipped upstream so we never reach here for them.
        if (!isVmergeContinue)
        {
            if (cell.Paragraphs.Count > 0)
            {
                Logger.Debug($"WriteTableCell: Writing {cell.Paragraphs.Count} paragraphs in cell");
                foreach (var para in cell.Paragraphs)
                {
                    Logger.Debug($"WriteTableCell: Calling WriteParagraph, Type={para.Type}, Text='{para.Text}', NestedTable={para.NestedTable != null}");
                    WriteParagraph(para);
                }
            }
            else
            {
                // OOXML requires at least one w:p in every w:tc
                _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteEndElement();
            }
        }
        else
        {
            // vMerge continue cells MUST still have at least one empty w:p per OOXML spec
            _writer.WriteStartElement("w", "p", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteEndElement();
        }
        
        _writer.WriteEndElement(); // w:tc
    }

    /// <summary>
    /// Returns true if the cell at (rowIndex, columnIndex) is within the vertical
    /// span of a cell above it (RowSpan &gt; 1).
    /// </summary>
    private static bool IsCoveredByVerticalMerge(TableModel table, int rowIndex, int columnIndex)
    {
        if (rowIndex <= 0) return false;

        for (int r = 0; r < rowIndex; r++)
        {
            var row = table.Rows[r];
            foreach (var c in row.Cells)
            {
                if (c.ColumnIndex != columnIndex) continue;
                if (c.RowSpan > 1)
                {
                    int start = c.RowIndex;
                    int end = c.RowIndex + c.RowSpan - 1;
                    if (rowIndex >= start && rowIndex <= end)
                    {
                        // This row is within the vertical span of the cell starting at 'start'
                        return rowIndex > start; // true for continuation rows only
                    }
                }
            }
        }

        return false;
    }
    
    /// <summary>
    /// Returns true if the cell at (rowIndex, columnIndex) is horizontally covered
    /// by a previous cell in the same row with ColumnSpan &gt; 1.
    /// </summary>
    private static bool IsCoveredByHorizontalMerge(TableModel table, int rowIndex, int columnIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count) return false;
        var row = table.Rows[rowIndex];
        if (columnIndex < 0 || columnIndex >= row.Cells.Count) return false;

        for (int c = 0; c < row.Cells.Count; c++)
        {
            var cell = row.Cells[c];
            if (cell.ColumnSpan > 1)
            {
                int spanStart = cell.ColumnIndex;
                int spanEnd = cell.ColumnIndex + cell.ColumnSpan - 1;
                if (columnIndex > spanStart && columnIndex <= spanEnd)
                {
                    return true;
                }
            }
        }

        return false;
    }
}
