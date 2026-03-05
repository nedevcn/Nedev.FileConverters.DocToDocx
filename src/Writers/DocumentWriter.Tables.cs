using System.Xml;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Writers;

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
        _writer.WriteStartElement("w", "tbl", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Write table properties
        WriteTableProperties(table);
        
        _writer.WriteStartElement("w", "tblGrid", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        int columnCount = table.ColumnCount > 0
            ? table.ColumnCount
            : (table.Rows.Any() ? table.Rows.Max(r => r.Cells.Count) : 0);
        for (int i = 0; i < columnCount; i++)
        {
            _writer.WriteStartElement("w", "gridCol", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            int width = 0;
            if (table.Rows.Count > 0 && i < table.Rows[0].Cells.Count && table.Rows[0].Cells[i].Properties?.Width > 0)
            {
                width = table.Rows[0].Cells[i].Properties!.Width;
            }
            
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
            WriteTableRow(row, table);
        }
        
        _writer.WriteEndElement(); // w:tbl
    }

    /// <summary>
    /// Writes table properties (tblPr).
    /// </summary>
    private void WriteTableProperties(TableModel table)
    {
        _writer.WriteStartElement("w", "tblPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Table style
        if (table.Properties?.StyleIndex >= 0)
        {
            var style = _document?.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Table && s.StyleId == table.Properties.StyleIndex);
            var styleId = StyleHelper.GetTableStyleId(table.Properties.StyleIndex, style?.Name);
            
            _writer.WriteStartElement("w", "tblStyle", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", styleId);
            _writer.WriteEndElement();
        }
        
        // Table width: prefer an explicit width from TAP when available, otherwise
        // let Word auto-size based on content.
        _writer.WriteStartElement("w", "tblW", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        var preferredWidth = table.Properties?.PreferredWidth ?? 0;
        if (preferredWidth > 0)
        {
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
        if (table.Properties != null && table.Properties.Alignment != TableAlignment.Left)
        {
            _writer.WriteStartElement("w", "jc", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            var alignment = table.Properties.Alignment switch
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
        if (table.Properties != null && table.Properties.Indent != 0)
        {
            _writer.WriteStartElement("w", "tblInd", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", table.Properties.Indent.ToString());
            _writer.WriteAttributeString("w", "type", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "dxa");
            _writer.WriteEndElement();
        }
        
        // Table borders
        if (table.Properties?.BorderTop != null || table.Properties?.BorderBottom != null ||
            table.Properties?.BorderLeft != null || table.Properties?.BorderRight != null ||
            table.Properties?.BorderInsideH != null || table.Properties?.BorderInsideV != null)
        {
            _writer.WriteStartElement("w", "tblBorders", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            if (table.Properties.BorderTop != null) WriteBorder("top", table.Properties.BorderTop);
            if (table.Properties.BorderBottom != null) WriteBorder("bottom", table.Properties.BorderBottom);
            if (table.Properties.BorderLeft != null) WriteBorder("left", table.Properties.BorderLeft);
            if (table.Properties.BorderRight != null) WriteBorder("right", table.Properties.BorderRight);
            if (table.Properties.BorderInsideH != null) WriteBorder("insideH", table.Properties.BorderInsideH);
            if (table.Properties.BorderInsideV != null) WriteBorder("insideV", table.Properties.BorderInsideV);
            _writer.WriteEndElement();
        }
        
        // Table shading
        if (table.Properties?.Shading != null)
        {
            WriteShading(table.Properties.Shading);
        }
        
        // Table cell margin: when the TAP exposes an inter-cell spacing we map it
        // to symmetric left/right padding; otherwise we fall back to a sensible
        // default that keeps existing documents visually similar.
        var spacing = table.Properties?.CellSpacing ?? 0;
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

    /// <summary>
    /// Writes a table row.
    /// </summary>
    private void WriteTableRow(TableRowModel row, TableModel table)
    {
        _writer.WriteStartElement("w", "tr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Row properties
        if (row.Properties != null)
        {
            _writer.WriteStartElement("w", "trPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            if (row.Properties.Height > 0)
            {
                _writer.WriteStartElement("w", "trHeight", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", row.Properties.Height.ToString());
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
            WriteTableCell(cell, row, table);
        }
        
        _writer.WriteEndElement(); // w:tr
    }

    /// <summary>
    /// Writes a table cell, including vertical (vMerge) and horizontal (gridSpan)
    /// merges. For vertical merges we emit w:vMerge restart/continue based on
    /// RowSpan and cells in previous rows; for horizontal merges we emit
    /// w:gridSpan on the first cell and suppress content in covered cells.
    /// </summary>
    private void WriteTableCell(TableCellModel cell, TableRowModel row, TableModel table)
    {
        _writer.WriteStartElement("w", "tc", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

        // Determine vertical merge role for this cell
        bool isVmergeStart = cell.RowSpan > 1;
        bool isVmergeContinue = !isVmergeStart && IsCoveredByVerticalMerge(table, row.Index, cell.ColumnIndex);
        
        bool hasTcPr = cell.Properties?.Width > 0 || cell.ColumnSpan > 1 || cell.RowSpan > 1 || isVmergeContinue ||
                       cell.Properties?.BorderTop != null || cell.Properties?.BorderBottom != null ||
                       cell.Properties?.BorderLeft != null || cell.Properties?.BorderRight != null ||
                       cell.Properties?.NoWrap == true ||
                       (cell.Properties != null && cell.Properties.VerticalAlignment != VerticalAlignment.Top);

        if (hasTcPr)
        {
            // tcPr: tcW -> gridSpan -> vMerge -> tcBorders -> shd -> noWrap -> vAlign
            _writer.WriteStartElement("w", "tcPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            
            // Cell width
            if (cell.Properties?.Width > 0)
            {
                _writer.WriteStartElement("w", "tcW", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                _writer.WriteAttributeString("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", cell.Properties.Width.ToString());
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
                foreach (var para in cell.Paragraphs)
                {
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
