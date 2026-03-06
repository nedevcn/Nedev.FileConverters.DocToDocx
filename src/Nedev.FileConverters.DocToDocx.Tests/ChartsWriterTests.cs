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
        [InlineData(ChartType.Doughnut, "pieChart")]
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
    }
}
