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
    }
}
