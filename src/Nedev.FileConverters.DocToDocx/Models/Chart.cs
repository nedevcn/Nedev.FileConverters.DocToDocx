namespace Nedev.FileConverters.DocToDocx.Models;

/// <summary>
/// Minimal chart model used to generate editable DOCX charts. This does not
/// attempt to mirror the full Excel/Office chart data model; instead it
/// captures enough structure to render a simple chart and to be a stable
/// target for future, richer parsing.
/// </summary>
public class ChartModel
{
    /// <summary>Zero-based chart index within the document.</summary>
    public int Index { get; set; }

    /// <summary>Optional title shown above the chart.</summary>
    public string? Title { get; set; }

    /// <summary>Chart type hint used by the writer to pick a CT_Chart subclass.</summary>
    public ChartType Type { get; set; } = ChartType.Column;

    /// <summary>Logical categories (X axis, or category axis for bar/column charts).</summary>
    public List<string> Categories { get; set; } = new();

    /// <summary>Data series that make up the chart.</summary>
    public List<ChartSeries> Series { get; set; } = new();

    /// <summary>
    /// Optional hint for where this chart should appear in the document
    /// (paragraph index). -1 means "no specific placement".
    /// </summary>
    public int ParagraphIndexHint { get; set; } = -1;

    /// <summary>
    /// Original OLE stream name from which this chart was detected
    /// (for example "Chart1"), when available.
    /// </summary>
    public string? SourceStreamName { get; set; }

    /// <summary>
    /// Raw bytes of the embedded chart container stream (usually the
    /// OLE/Excel workbook or MSGraph chart stream). This allows future
    /// phases or external tools to perform deeper parsing (e.g. BIFF,
    /// Excel, MSGraph) without having to re‑open the original `.doc`.
    /// May be null when the stream could not be read.
    /// </summary>
    public byte[]? SourceBytes { get; set; }

    /// <summary>
    /// Convenience alias of <see cref="SourceBytes"/> that emphasises the
    /// fact that the payload is typically the original workbook bytes.
    /// Both properties refer to the same underlying array; setting one
    /// updates the other.
    /// </summary>
    public byte[]? WorkbookBytes
    {
        get => SourceBytes;
        set => SourceBytes = value;
    }

    /// <summary>
    /// Optional color of the chart title text (RGB integer, zero = none).
    /// </summary>
    public int TitleColor { get; set; }

    /// <summary>
    /// Optional title displayed on the category (X) axis.
    /// </summary>
    public string? CategoryAxisTitle { get; set; }

    /// <summary>
    /// Optional color for the category-axis title text (RGB integer, zero = none).
    /// </summary>
    public int CategoryAxisTitleColor { get; set; }

    /// <summary>
    /// Optional color for the value-axis title text (RGB integer, zero = none).
    /// </summary>
    public int ValueAxisTitleColor { get; set; }

    /// <summary>
    /// Optional title displayed on the value (Y) axis.
    /// </summary>
    public string? ValueAxisTitle { get; set; }

    /// <summary>
    /// Whether a legend should be shown for the chart. Defaults to true.
    /// </summary>
    public bool ShowLegend { get; set; } = true;
}

public enum ChartType
{
    Column,
    Line,
    Bar,
    Pie,
    Area,
    Scatter,
    Doughnut,
    Radar,
    Unknown
}

/// <summary>
/// One data series in a chart (e.g. "Sales 2024").
/// </summary>
public class ChartSeries
{
    public string Name { get; set; } = string.Empty;

    /// <summary>
    /// Values aligned with the parent ChartModel.Categories collection.
    /// When the value count does not match the number of categories, the
    /// writer will truncate or pad with zeros as needed.
    /// </summary>
    public List<double> Values { get; set; } = new();
}

