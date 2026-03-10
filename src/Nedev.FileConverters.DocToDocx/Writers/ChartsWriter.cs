using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

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

    /// <summary>
    /// Creates a new <see cref="ChartsWriter"/> that will emit chart XML to the
    /// supplied <see cref="XmlWriter"/>.
    /// </summary>
    /// <param name="writer">The XML writer to which chart parts will be written.</param>
    public ChartsWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    /// <summary>
    /// Writes a <c>chartN.xml</c> part representing the provided <see cref="ChartModel"/>.
    /// The generated XML is intentionally sparse, focusing on producing a valid,
    /// editable chart in Word rather than reproducing every legacy formatting detail.
    /// </summary>
    /// <param name="chart">The chart model to serialize.</param>
    public void WriteChart(ChartModel chart)
    {
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        _writer.WriteStartDocument();
        _writer.WriteStartElement("c", "chartSpace", cNs);

        _writer.WriteAttributeString("xmlns", "c", null, cNs);
        _writer.WriteAttributeString("xmlns", "a", null, aNs);
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        _writer.WriteStartElement("c", "lang", cNs);
        _writer.WriteAttributeString("val", "en-US");
        _writer.WriteEndElement();

        _writer.WriteStartElement("c", "roundedCorners", cNs);
        _writer.WriteAttributeString("val", "0");
        _writer.WriteEndElement();

        _writer.WriteStartElement("c", "chart", cNs);

        if (!string.IsNullOrEmpty(chart.Title))
        {
            WriteRichTextElement("title", chart.Title, cNs, aNs, chart.TitleColor);
        }
        else
        {
            _writer.WriteStartElement("c", "autoTitleDeleted", cNs);
            _writer.WriteAttributeString("val", "1");
            _writer.WriteEndElement();
        }

        _writer.WriteStartElement("c", "plotArea", cNs);
        _writer.WriteStartElement("c", GetChartElementName(chart.Type), cNs);

        bool isPie = chart.Type == ChartType.Pie || chart.Type == ChartType.Doughnut;
        WriteChartTypeOptions(chart, cNs);
        if (!isPie)
        {
            WriteCategoryAxisData(chart);
        }

        WriteSeriesData(chart);

        if (!isPie)
        {
            WriteAxisReferences(cNs);
        }

        _writer.WriteEndElement();

        if (!isPie)
        {
            WriteDefaultAxes(chart);
        }

        _writer.WriteEndElement();

        if (chart.ShowLegend)
        {
            _writer.WriteStartElement("c", "legend", cNs);
            _writer.WriteStartElement("c", "legendPos", cNs);
            _writer.WriteAttributeString("val", "r");
            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }

        _writer.WriteStartElement("c", "plotVisOnly", cNs);
        _writer.WriteAttributeString("val", "1");
        _writer.WriteEndElement();

        _writer.WriteStartElement("c", "dispBlanksAs", cNs);
        _writer.WriteAttributeString("val", "gap");
        _writer.WriteEndElement();

        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }

    /// <summary>
    /// Gets the root chart element name for a given <see cref="ChartType"/>, e.g.
    /// <c>barChart</c> or <c>lineChart</c>.
    /// </summary>
    /// <param name="type">The generic chart type.</param>
    /// <returns>The string name used in the OOXML schema.</returns>
    public static string GetChartElementName(ChartType type) => type switch
    {
        ChartType.Line => "lineChart",
        ChartType.Bar => "barChart",
        ChartType.Column => "barChart",
        ChartType.Pie => "pieChart",
        ChartType.Doughnut => "doughnutChart",
        ChartType.Area => "areaChart",
        ChartType.Scatter => "scatterChart",
        ChartType.Radar => "radarChart",
        _ => "barChart"
    };

    private void WriteDefaultAxes(ChartModel chart)
    {
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        const string aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

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
        _writer.WriteStartElement("c", "delete", cNs);
        _writer.WriteAttributeString("val", "0");
        _writer.WriteEndElement();
        if (!string.IsNullOrEmpty(chart.CategoryAxisTitle))
        {
            WriteRichTextElement("title", chart.CategoryAxisTitle, cNs, aNs, chart.CategoryAxisTitleColor);
        }
        _writer.WriteStartElement("c", "tickLblPos", cNs);
        _writer.WriteAttributeString("val", "nextTo");
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "lblOffset", cNs);
        _writer.WriteAttributeString("val", "100");
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "crossAx", cNs);
        _writer.WriteAttributeString("val", "2");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

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
        _writer.WriteStartElement("c", "delete", cNs);
        _writer.WriteAttributeString("val", "0");
        _writer.WriteEndElement();
        if (!string.IsNullOrEmpty(chart.ValueAxisTitle))
        {
            WriteRichTextElement("title", chart.ValueAxisTitle, cNs, aNs, chart.ValueAxisTitleColor);
        }
        _writer.WriteStartElement("c", "majorGridlines", cNs);
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "tickLblPos", cNs);
        _writer.WriteAttributeString("val", "nextTo");
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "crosses", cNs);
        _writer.WriteAttributeString("val", "autoZero");
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "crossBetween", cNs);
        _writer.WriteAttributeString("val", "between");
        _writer.WriteEndElement();
        _writer.WriteStartElement("c", "crossAx", cNs);
        _writer.WriteAttributeString("val", "1");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }

    private void WriteChartTypeOptions(ChartModel chart, string cNs)
    {
        switch (chart.Type)
        {
            case ChartType.Bar:
            case ChartType.Column:
                WriteBoolValElement("c", "varyColors", cNs, false);
                _writer.WriteStartElement("c", "barDir", cNs);
                _writer.WriteAttributeString("val", chart.Type == ChartType.Bar ? "bar" : "col");
                _writer.WriteEndElement();
                _writer.WriteStartElement("c", "grouping", cNs);
                _writer.WriteAttributeString("val", "clustered");
                _writer.WriteEndElement();
                _writer.WriteStartElement("c", "gapWidth", cNs);
                _writer.WriteAttributeString("val", "150");
                _writer.WriteEndElement();
                _writer.WriteStartElement("c", "overlap", cNs);
                _writer.WriteAttributeString("val", "0");
                _writer.WriteEndElement();
                break;
            case ChartType.Line:
                WriteBoolValElement("c", "varyColors", cNs, false);
                _writer.WriteStartElement("c", "grouping", cNs);
                _writer.WriteAttributeString("val", "standard");
                _writer.WriteEndElement();
                WriteBoolValElement("c", "marker", cNs, false);
                WriteBoolValElement("c", "smooth", cNs, false);
                break;
            case ChartType.Area:
                WriteBoolValElement("c", "varyColors", cNs, false);
                _writer.WriteStartElement("c", "grouping", cNs);
                _writer.WriteAttributeString("val", "standard");
                _writer.WriteEndElement();
                break;
            case ChartType.Scatter:
                WriteBoolValElement("c", "varyColors", cNs, false);
                _writer.WriteStartElement("c", "scatterStyle", cNs);
                _writer.WriteAttributeString("val", "lineMarker");
                _writer.WriteEndElement();
                break;
            case ChartType.Radar:
                WriteBoolValElement("c", "varyColors", cNs, false);
                _writer.WriteStartElement("c", "radarStyle", cNs);
                _writer.WriteAttributeString("val", "standard");
                _writer.WriteEndElement();
                break;
            case ChartType.Pie:
            case ChartType.Doughnut:
                WriteBoolValElement("c", "varyColors", cNs, true);
                _writer.WriteStartElement("c", "firstSliceAng", cNs);
                _writer.WriteAttributeString("val", "0");
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

    private void WriteAxisReferences(string cNs)
    {
        _writer.WriteStartElement("c", "axId", cNs);
        _writer.WriteAttributeString("val", "1");
        _writer.WriteEndElement();

        _writer.WriteStartElement("c", "axId", cNs);
        _writer.WriteAttributeString("val", "2");
        _writer.WriteEndElement();
    }

    private void WriteRichTextElement(string elementName, string text, string cNs, string aNs, int color = 0)
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
        if (color != 0)
        {
            _writer.WriteStartElement("a", "rPr", aNs);
            _writer.WriteStartElement("a", "solidFill", aNs);
            _writer.WriteStartElement("a", "srgbClr", aNs);
            _writer.WriteAttributeString("val", ColorHelper.ResolveColorHex(color, null, "000000"));
            _writer.WriteEndElement(); // a:srgbClr
            _writer.WriteEndElement(); // a:solidFill
            _writer.WriteEndElement(); // a:rPr
        }
        WriteTextElement("a", "t", aNs, text);
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }

    private void WriteCategoryAxisData(ChartModel chart)
    {
        var categories = chart.Categories;
        if (categories.Count == 0)
        {
            categories = new List<string> { "Category 1", "Category 2", "Category 3" };
        }
    }

    private void WriteSeriesData(ChartModel chart)
    {
        const string cNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        var categories = chart.Categories.Count > 0
            ? chart.Categories
            : new List<string> { "Category 1", "Category 2", "Category 3" };

        IReadOnlyList<ChartSeries> seriesList = chart.Series.Count > 0
            ? chart.Series
            : new[]
            {
                new ChartSeries
                {
                    Name = "Series 1",
                    Values = new List<double> { 1, 2, 3 }
                }
            };

        for (int s = 0; s < seriesList.Count; s++)
        {
            var series = seriesList[s];
            var seriesName = !string.IsNullOrWhiteSpace(series.Name) ? series.Name : $"Series {s + 1}";
            _writer.WriteStartElement("c", "ser", cNs);

            _writer.WriteStartElement("c", "idx", cNs);
            _writer.WriteAttributeString("val", s.ToString());
            _writer.WriteEndElement();
            _writer.WriteStartElement("c", "order", cNs);
            _writer.WriteAttributeString("val", s.ToString());
            _writer.WriteEndElement();

            _writer.WriteStartElement("c", "tx", cNs);
            _writer.WriteStartElement("c", "strRef", cNs);
            _writer.WriteStartElement("c", "strCache", cNs);
            _writer.WriteStartElement("c", "ptCount", cNs);
            _writer.WriteAttributeString("val", "1");
            _writer.WriteEndElement();
            _writer.WriteStartElement("c", "pt", cNs);
            _writer.WriteAttributeString("idx", "0");
            WriteTextElement("c", "v", cNs, seriesName);
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            WriteBoolValElement("c", "invertIfNegative", cNs, false);

            _writer.WriteStartElement("c", "cat", cNs);
            _writer.WriteStartElement("c", "strRef", cNs);
            _writer.WriteStartElement("c", "strCache", cNs);
            _writer.WriteStartElement("c", "ptCount", cNs);
            _writer.WriteAttributeString("val", categories.Count.ToString());
            _writer.WriteEndElement();
            for (int i = 0; i < categories.Count; i++)
            {
                _writer.WriteStartElement("c", "pt", cNs);
                _writer.WriteAttributeString("idx", i.ToString());
                WriteTextElement("c", "v", cNs, categories[i]);
                _writer.WriteEndElement();
            }
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.WriteStartElement("c", "val", cNs);
            _writer.WriteStartElement("c", "numRef", cNs);
            _writer.WriteStartElement("c", "numCache", cNs);

            var values = series.Values;
            if (values.Count < categories.Count)
            {
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
            _writer.WriteEndElement();

            for (int i = 0; i < values.Count; i++)
            {
                _writer.WriteStartElement("c", "pt", cNs);
                _writer.WriteAttributeString("idx", i.ToString());
                _writer.WriteStartElement("c", "v", cNs);
                _writer.WriteString(values[i].ToString(System.Globalization.CultureInfo.InvariantCulture));
                _writer.WriteEndElement();
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement();

            _writer.WriteEndElement();
        }
    }

    private void WriteTextElement(string prefix, string localName, string ns, string? text)
    {
        var safeText = DocumentWriter.SanitizeXmlString(text ?? string.Empty);
        _writer.WriteStartElement(prefix, localName, ns);
        if (NeedsSpacePreserve(safeText))
        {
            _writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        }
        _writer.WriteString(safeText);
        _writer.WriteEndElement();
    }

    private void WriteBoolValElement(string prefix, string localName, string ns, bool value)
    {
        _writer.WriteStartElement(prefix, localName, ns);
        _writer.WriteAttributeString("val", value ? "1" : "0");
        _writer.WriteEndElement();
    }

    private static bool NeedsSpacePreserve(string text)
    {
        return !string.IsNullOrEmpty(text)
            && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]) || text.Contains("  "));
    }
}

