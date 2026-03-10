using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// Writes minimal but standards-compliant DOCX chart parts (chartN.xml) from
/// the high-level ChartModel. The goal is to produce charts that open and are
/// editable in Word, even when the underlying .doc chart data could not be
/// fully recovered.
/// </summary>
public class ChartsWriter
{
    private readonly XmlWriter _writer;

    public ChartsWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    public void WriteChart(ChartModel chart)
    {
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        _writer.WriteStartDocument();
        _writer.WriteStartElement("c", "chartSpace", cNs);

        // Namespaces
        _writer.WriteAttributeString("xmlns", "c", null, cNs);
        _writer.WriteAttributeString("xmlns", "a", null, aNs);
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        _writer.WriteStartElement("c", "lang", cNs);
        _writer.WriteAttributeString("val", "en-US");
        _writer.WriteEndElement();

        // Chart properties (very minimal)
        _writer.WriteStartElement("c", "chart", cNs);

        // Title (optional)
        if (!string.IsNullOrEmpty(chart.Title))
        {
            WriteRichTextElement("title", chart.Title, cNs, aNs);
        }

        // Plot area with a single chart type. For most chart types we use a
        // simple category/value axis layout. Pie‑like charts do not require axes
        // or cat/val elements, but the series data itself is still written.
        _writer.WriteStartElement("c", "plotArea", cNs);
        _writer.WriteStartElement("c", GetChartElementName(chart.Type), cNs);

        bool isPie = chart.Type == ChartType.Pie || chart.Type == ChartType.Doughnut;
        WriteChartTypeOptions(chart, cNs);
        if (!isPie)
        {
            WriteCategoryAxisData(chart);
        }

        WriteSeriesData(chart);

        _writer.WriteEndElement(); // c:chartType

        // Axes (catAx + valAx) with default ids; omit for pie/doughnut
        if (!isPie)
        {
            WriteDefaultAxes(chart);
        }

        _writer.WriteEndElement(); // c:plotArea

        // Legend (optional simple right-side legend)
        if (chart.ShowLegend)
        {
            _writer.WriteStartElement("c", "legend", cNs);
            _writer.WriteStartElement("c", "legendPos", cNs);
            _writer.WriteAttributeString("val", "r");
            _writer.WriteEndElement(); // c:legendPos
            _writer.WriteEndElement(); // c:legend
        }

        _writer.WriteEndElement(); // c:chart
        _writer.WriteEndElement(); // c:chartSpace
        _writer.WriteEndDocument();
    }

    // made internal so that unit tests can verify mapping without needing to
    // drive the entire writer.
    public static string GetChartElementName(ChartType type) => type switch
    {
        ChartType.Line => "lineChart",
        ChartType.Bar => "barChart",
        ChartType.Column => "barChart", // column is semantically a clustered bar in CT
        ChartType.Pie => "pieChart",
        ChartType.Doughnut => "doughnutChart",
        ChartType.Area => "areaChart",
        ChartType.Scatter => "scatterChart",
        ChartType.Radar => "radarChart",
        _ => "barChart" // fallback to something reasonable
    };

    /// <summary>
    /// Writes c:catAx and c:valAx with fixed ids (1 and 2). This is enough for
    /// Word to treat the part as a valid chart with category/value axes.
    /// </summary>
    private void WriteDefaultAxes(ChartModel chart)
    {
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        // Category axis
        _writer.WriteStartElement("c", "catAx", cNs);
        _writer.WriteStartElement("c", "axId", cNs);
        _writer.WriteAttributeString("val", "1");
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "scaling", cNs);
        _writer.WriteStartElement("c", "orientation", cNs);
        _writer.WriteAttributeString("val", "minMax");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "axPos", cNs);
        _writer.WriteAttributeString("val", "b");
        _writer.WriteEndElement();
        if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
        {
            WriteRichTextElement("title", chart.CategoryAxisTitle, cNs, aNs);
        }
        _writer.WriteStartElement("c", "crossAx", cNs);
        _writer.WriteAttributeString("val", "2");
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // c:catAx

        // Value axis
        _writer.WriteStartElement("c", "valAx", cNs);
        _writer.WriteStartElement("c", "axId", cNs);
        _writer.WriteAttributeString("val", "2");
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "scaling", cNs);
        _writer.WriteStartElement("c", "orientation", cNs);
        _writer.WriteAttributeString("val", "minMax");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "axPos", cNs);
        _writer.WriteAttributeString("val", "l");
        _writer.WriteEndElement();
        if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
        {
            WriteRichTextElement("title", chart.ValueAxisTitle, cNs, aNs);
        }
        _writer.WriteStartElement("c", "crossAx", cNs);
        _writer.WriteAttributeString("val", "1");
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // c:valAx
    }

    private void WriteChartTypeOptions(ChartModel chart, string cNs)
    {
        switch (chart.Type)
        {
            case ChartType.Bar:
            case ChartType.Column:
                _writer.WriteStartElement("c", "barDir", cNs);
                _writer.WriteAttributeString("val", chart.Type == ChartType.Bar ? "bar" : "col");
                _writer.WriteEndElement();
                break;
            case ChartType.Line:
            case ChartType.Area:
                _writer.WriteStartElement("c", "grouping", cNs);
                _writer.WriteAttributeString("val", "standard");
                _writer.WriteEndElement();
                break;
            case ChartType.Scatter:
                _writer.WriteStartElement("c", "scatterStyle", cNs);
                _writer.WriteAttributeString("val", "lineMarker");
                _writer.WriteEndElement();
                break;
            case ChartType.Radar:
                _writer.WriteStartElement("c", "radarStyle", cNs);
                _writer.WriteAttributeString("val", "standard");
                _writer.WriteEndElement();
                break;
            case ChartType.Pie:
            case ChartType.Doughnut:
                _writer.WriteStartElement("c", "varyColors", cNs);
                _writer.WriteAttributeString("val", "1");
                _writer.WriteEndElement();
                if (chart.Type == ChartType.Doughnut)
                {
                    _writer.WriteStartElement("c", "holeSize", cNs);
                    _writer.WriteAttributeString("val", "50");
                    _writer.WriteEndElement();
                }
                break;
        }
    }

    private void WriteRichTextElement(string elementName, string text, string cNs, string aNs)
    {
        _writer.WriteStartElement("c", elementName, cNs);
        _writer.WriteStartElement("c", "tx", cNs);
        _writer.WriteStartElement("c", "rich", cNs);
        _writer.WriteStartElement("a", "bodyPr", aNs);
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "lstStyle", aNs);
        _writer.WriteEndElement();
        _writer.WriteStartElement("a", "p", aNs);
        _writer.WriteStartElement("a", "r", aNs);
        _writer.WriteStartElement("a", "t", aNs);
        _writer.WriteString(text);
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }

    /// <summary>
    /// Writes a minimal c:cat element with inline string categories.
    /// </summary>
    private void WriteCategoryAxisData(ChartModel chart)
    {
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        var categories = chart.Categories;
        if (categories.Count == 0)
        {
            categories = new List<string> { "Category 1", "Category 2", "Category 3" };
        }

        // Category axis data is written per series in c:cat, but when using
        // inline data it is sufficient to emit s:pt entries alongside series.
        // We handle per-series cats in WriteSeriesData, so nothing to do here.
    }

    /// <summary>
    /// Writes all series and their values. We emit inline categories/values
    /// rather than external references to a worksheet.
    /// </summary>
    private void WriteSeriesData(ChartModel chart)
    {
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        var categories = chart.Categories.Count > 0
            ? chart.Categories
            : new List<string> { "Category 1", "Category 2", "Category 3" };

        if (chart.Series.Count == 0)
        {
            chart.Series.Add(new ChartSeries
            {
                Name = "Series 1",
                Values = new List<double> { 1, 2, 3 }
            });
        }

        for (int s = 0; s < chart.Series.Count; s++)
        {
            var series = chart.Series[s];
            _writer.WriteStartElement("c", "ser", cNs);

            // Index and order
            _writer.WriteStartElement("c", "idx", cNs);
            _writer.WriteAttributeString("val", s.ToString());
            _writer.WriteEndElement();
            _writer.WriteStartElement("c", "order", cNs);
            _writer.WriteAttributeString("val", s.ToString());
            _writer.WriteEndElement();

            // Series name
            if (!string.IsNullOrEmpty(series.Name))
            {
                _writer.WriteStartElement("c", "tx", cNs);
                _writer.WriteStartElement("c", "strRef", cNs);
                _writer.WriteStartElement("c", "strCache", cNs);
                _writer.WriteStartElement("c", "ptCount", cNs);
                _writer.WriteAttributeString("val", "1");
                _writer.WriteEndElement(); // c:ptCount
                _writer.WriteStartElement("c", "pt", cNs);
                _writer.WriteAttributeString("idx", "0");
                _writer.WriteStartElement("c", "v", cNs);
                _writer.WriteString(series.Name);
                _writer.WriteEndElement(); // c:v
                _writer.WriteEndElement(); // c:pt
                _writer.WriteEndElement(); // c:strCache
                _writer.WriteEndElement(); // c:strRef
                _writer.WriteEndElement(); // c:tx
            }

            // Categories (string cache)
            _writer.WriteStartElement("c", "cat", cNs);
            _writer.WriteStartElement("c", "strRef", cNs);
            _writer.WriteStartElement("c", "strCache", cNs);
            _writer.WriteStartElement("c", "ptCount", cNs);
            _writer.WriteAttributeString("val", categories.Count.ToString());
            _writer.WriteEndElement(); // c:ptCount
            for (int i = 0; i < categories.Count; i++)
            {
                _writer.WriteStartElement("c", "pt", cNs);
                _writer.WriteAttributeString("idx", i.ToString());
                _writer.WriteStartElement("c", "v", cNs);
                _writer.WriteString(categories[i]);
                _writer.WriteEndElement(); // c:v
                _writer.WriteEndElement(); // c:pt
            }
            _writer.WriteEndElement(); // c:strCache
            _writer.WriteEndElement(); // c:strRef
            _writer.WriteEndElement(); // c:cat

            // Values (number cache)
            _writer.WriteStartElement("c", "val", cNs);
            _writer.WriteStartElement("c", "numRef", cNs);
            _writer.WriteStartElement("c", "numCache", cNs);

            // Align value count with categories
            var values = series.Values;
            if (values.Count < categories.Count)
            {
                // pad with zeros
                var padded = new List<double>(values);
                while (padded.Count < categories.Count)
                    padded.Add(0);
                values = padded;
            }
            else if (values.Count > categories.Count)
            {
                values = values.Take(categories.Count).ToList();
            }

            _writer.WriteStartElement("c", "ptCount", cNs);
            _writer.WriteAttributeString("val", values.Count.ToString());
            _writer.WriteEndElement(); // c:ptCount

            for (int i = 0; i < values.Count; i++)
            {
                _writer.WriteStartElement("c", "pt", cNs);
                _writer.WriteAttributeString("idx", i.ToString());
                _writer.WriteStartElement("c", "v", cNs);
                _writer.WriteString(values[i].ToString(System.Globalization.CultureInfo.InvariantCulture));
                _writer.WriteEndElement(); // c:v
                _writer.WriteEndElement(); // c:pt
            }

            _writer.WriteEndElement(); // c:numCache
            _writer.WriteEndElement(); // c:numRef
            _writer.WriteEndElement(); // c:val

            _writer.WriteEndElement(); // c:ser
        }
    }
}

