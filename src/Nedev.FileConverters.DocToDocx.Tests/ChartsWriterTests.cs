#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Writers;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests
{
    public class ChartsWriterTests
    {
        [Theory]
        [InlineData(ChartType.Column, "barChart")]
        [InlineData(ChartType.Line, "lineChart")]
        [InlineData(ChartType.Bar, "barChart")]
        [InlineData(ChartType.Pie, "pieChart")]
        [InlineData(ChartType.Doughnut, "doughnutChart")]
        [InlineData(ChartType.Area, "areaChart")]
        [InlineData(ChartType.Scatter, "scatterChart")]
        [InlineData(ChartType.Radar, "radarChart")]
        [InlineData(ChartType.Unknown, "barChart")]
        public void GetChartElementName_ReturnsCorrectMapping(ChartType type, string expected)
        {
            var actual = ChartsWriter.GetChartElementName(type);
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void WriteChart_Pie_DoesNotEmitAxes()
        {
            var model = new ChartModel
            {
                Type = ChartType.Pie,
                Categories = new List<string> { "A", "B" },
                Series = { new ChartSeries { Name = "S", Values = new List<double> { 1, 2 } } }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new System.Xml.XmlWriterSettings { Encoding = Encoding.UTF8 });
            var cw = new ChartsWriter(writer);
            cw.WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            Assert.Contains("<c:pieChart", xml);
            Assert.DoesNotContain("<c:catAx", xml);
            Assert.DoesNotContain("<c:valAx", xml);
        }

        [Fact]
        public void WriteChart_EmitsLegendAndAxisTitlesWhenProvided()
        {
            var model = new ChartModel
            {
                Type = ChartType.Column,
                Title = "Quarterly Sales",
                CategoryAxisTitle = "Quarter",
                ValueAxisTitle = "Revenue",
                ShowLegend = true,
                Categories = new List<string> { "Q1", "Q2" },
                Series =
                {
                    new ChartSeries { Name = "North", Values = new List<double> { 10, 12 } },
                    new ChartSeries { Name = "South", Values = new List<double> { 11, 13 } }
                }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer).WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            Assert.Contains("Quarterly Sales", xml);
            Assert.Contains("Quarter", xml);
            Assert.Contains("Revenue", xml);
            Assert.Contains("<c:legend>", xml);
            Assert.Contains("<c:barDir val=\"col\"", xml);
        }

        [Fact]
        public void WriteChart_Doughnut_EmitsHoleSizeAndSkipsLegendWhenDisabled()
        {
            var model = new ChartModel
            {
                Type = ChartType.Doughnut,
                ShowLegend = false,
                Categories = new List<string> { "A", "B" },
                Series = { new ChartSeries { Name = "S", Values = new List<double> { 1, 2 } } }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer).WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            Assert.Contains("<c:doughnutChart", xml);
            Assert.Contains("<c:firstSliceAng val=\"0\"", xml);
            Assert.Contains("<c:holeSize val=\"50\"", xml);
            Assert.DoesNotContain("<c:legend>", xml);
        }

        [Fact]
        public void WriteChart_AxisBasedCharts_EmitAxisReferencesAndDisplayDefaults()
        {
            var model = new ChartModel
            {
                Type = ChartType.Column,
                Categories = new List<string> { "A", "B" },
                Series = { new ChartSeries { Name = "S", Values = new List<double> { 1, 2 } } }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer).WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            Assert.Contains("<c:autoTitleDeleted val=\"1\"", xml);
            Assert.Contains("<c:plotVisOnly val=\"1\"", xml);
            Assert.Contains("<c:dispBlanksAs val=\"gap\"", xml);
            Assert.True(xml.Split("<c:axId val=\"1\"", StringSplitOptions.None).Length >= 3);
            Assert.True(xml.Split("<c:axId val=\"2\"", StringSplitOptions.None).Length >= 3);
        }

        [Fact]
        public void WriteChart_BarChart_EmitsClusteredAndAxisDefaults()
        {
            var model = new ChartModel
            {
                Type = ChartType.Bar,
                Categories = new List<string> { "A", "B" },
                Series = { new ChartSeries { Name = "S", Values = new List<double> { 1, 2 } } }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer).WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            Assert.Contains("<c:roundedCorners val=\"0\"", xml);
            Assert.Contains("<c:varyColors val=\"0\"", xml);
            Assert.Contains("<c:grouping val=\"clustered\"", xml);
            Assert.Contains("<c:gapWidth val=\"150\"", xml);
            Assert.Contains("<c:overlap val=\"0\"", xml);
            Assert.Contains("<c:delete val=\"0\"", xml);
            Assert.Contains("<c:tickLblPos val=\"nextTo\"", xml);
            Assert.Contains("<c:lblOffset val=\"100\"", xml);
            Assert.Contains("<c:majorGridlines", xml);
            Assert.Contains("<c:crosses val=\"autoZero\"", xml);
            Assert.Contains("<c:crossBetween val=\"between\"", xml);
        }

        [Fact]
        public void WriteChart_LineChart_SanitizesTextAndEmitsLineDefaults()
        {
            var model = new ChartModel
            {
                Type = ChartType.Line,
                Title = " Sales\u0001 Report ",
                CategoryAxisTitle = " Quarter\u0001 ",
                ValueAxisTitle = " Revenue\uFFFD ",
                Categories = new List<string> { " Q1 ", "Q2\u0001" },
                Series = { new ChartSeries { Name = " North\u0001 ", Values = new List<double> { 1, 2 } } }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer).WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            Assert.Contains("<c:marker val=\"0\"", xml);
            Assert.Contains("<c:smooth val=\"0\"", xml);
            Assert.Contains("xml:space=\"preserve\"", xml);
            Assert.Contains(" Sales Report ", xml);
            Assert.Contains(" Quarter ", xml);
            Assert.Contains(" Revenue  ", xml);
            Assert.Contains(" North ", xml);
            Assert.Contains(" Q1 ", xml);

            using var reader = XmlReader.Create(new StringReader(xml.TrimStart('\uFEFF')));
            while (reader.Read()) { }
        }

        [Fact]
        public void WriteChart_DoesNotMutateSourceSeries_AndFallsBackForBlankSeriesNames()
        {
            var emptyModel = new ChartModel
            {
                Type = ChartType.Column,
                Categories = new List<string> { "A", "B", "C" }
            };

            using (var ms = new MemoryStream())
            using (var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 }))
            {
                new ChartsWriter(writer).WriteChart(emptyModel);
                writer.Flush();
            }

            Assert.Empty(emptyModel.Series);

            var blankNameModel = new ChartModel
            {
                Type = ChartType.Column,
                Categories = new List<string> { "A" },
                Series = { new ChartSeries { Name = " ", Values = new List<double> { 1 } } }
            };

            using var ms2 = new MemoryStream();
            using var writer2 = XmlWriter.Create(ms2, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer2).WriteChart(blankNameModel);
            writer2.Flush();
            string xml = Encoding.UTF8.GetString(ms2.ToArray());

            Assert.Contains(">Series 1</c:v>", xml);
        }

        [Fact]
        public void WriteChart_SeriesValueMismatch_PadsAndTruncates()
        {
            var model = new ChartModel
            {
                Type = ChartType.Column,
                Categories = new List<string> { "A", "B", "C" },
                Series =
                {
                    new ChartSeries { Name = "S1", Values = new List<double> { 1, 2 } },
                    new ChartSeries { Name = "S2", Values = new List<double> { 3, 4, 5, 6 } }
                }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer).WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            // first series should be padded to three values, second truncated to three
            Assert.Contains("<c:v>1</c:v>", xml);
            Assert.Contains("<c:v>2</c:v>", xml);
            Assert.Contains("<c:v>0</c:v>", xml); // padding
            Assert.DoesNotContain("<c:v>6</c:v>", xml); // truncated
        }

        [Fact]
        public void WriteChart_RespectsTextColors()
        {
            var model = new ChartModel
            {
                Type = ChartType.Column,
                Title = "Colorful",
                CategoryAxisTitle = "X",
                ValueAxisTitle = "Y",
                TitleColor = 0x112233,
                CategoryAxisTitleColor = 0x445566,
                ValueAxisTitleColor = 0x778899,
                Categories = new List<string> { "A" },
                Series = { new ChartSeries { Name = "S", Values = new List<double> { 1 } } }
            };

            using var ms = new MemoryStream();
            using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
            new ChartsWriter(writer).WriteChart(model);
            writer.Flush();
            string xml = Encoding.UTF8.GetString(ms.ToArray());

            // ColorHelper outputs BGR hex, so the byte order is reversed compared to the numeric value.
            Assert.Contains("srgbClr val=\"332211\"", xml);
            Assert.Contains("srgbClr val=\"665544\"", xml);
            Assert.Contains("srgbClr val=\"998877\"", xml);
        }
    }
}
