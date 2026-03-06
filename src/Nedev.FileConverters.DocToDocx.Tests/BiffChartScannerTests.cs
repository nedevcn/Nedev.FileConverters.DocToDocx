#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests
{
    public class BiffChartScannerTests
    {
        private static byte[] BuildBiffMatrix(double[,] values, string[]? rowLabels = null, string[]? colLabels = null)
        {
            using var ms = new MemoryStream();
            using var w = new BinaryWriter(ms);
            // minimal BOF record so TryPopulateChart recognizes stream
            w.Write((ushort)0x0809);
            w.Write((ushort)0);

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            // emit numeric cells
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var v = values[r, c];
                    if (!double.IsNaN(v))
                    {
                        w.Write((ushort)0x0203);
                        w.Write((ushort)14);
                        w.Write((ushort)r);
                        w.Write((ushort)c);
                        w.Write((ushort)0);
                        w.Write(v);
                    }
                }
            }

            if (rowLabels != null)
            {
                for (int r = 0; r < rowLabels.Length; r++)
                {
                    var s = rowLabels[r];
                    var bytes = Encoding.Default.GetBytes(s);
                    w.Write((ushort)0x0204);
                    w.Write((ushort)(8 + 1 + bytes.Length));
                    w.Write((ushort)(r + 1));
                    w.Write((ushort)0);
                    w.Write((ushort)0);
                    w.Write((ushort)bytes.Length);
                    w.Write((byte)0);
                    w.Write(bytes);
                }
            }

            if (colLabels != null)
            {
                for (int c = 0; c < colLabels.Length; c++)
                {
                    var s = colLabels[c];
                    var bytes = Encoding.Default.GetBytes(s);
                    w.Write((ushort)0x0204);
                    w.Write((ushort)(8 + 1 + bytes.Length));
                    w.Write((ushort)0);
                    w.Write((ushort)(c + 1));
                    w.Write((ushort)0);
                    w.Write((ushort)bytes.Length);
                    w.Write((byte)0);
                    w.Write(bytes);
                }
            }

            return ms.ToArray();
        }

        [Fact]
        public void ParsesSimpleTable_ProducesExpectedSeriesAndCategories()
        {
            // layout: header row with 2 category names, first column as series name
            double[,] data = {
                { double.NaN, 10, 20 },
                { 0, 1, 2 },
                { 0, 3, 4 }
            };
            var bytes = BuildBiffMatrix(data, rowLabels: new[] { "Series1", "Series2" }, colLabels: new[] { "Cat1", "Cat2" });

            var model = new ChartModel { SourceBytes = bytes, SourceStreamName = "Chart1" };
            BiffChartScanner.TryPopulateChart(model);

            Assert.Equal(2, model.Categories.Count);
            Assert.Equal("Cat1", model.Categories[0]);
            Assert.Equal("Cat2", model.Categories[1]);
            Assert.Equal(2, model.Series.Count);
            Assert.Equal("Series1", model.Series[0].Name);
            Assert.Equal(new List<double> { 1, 2 }, model.Series[0].Values);
            Assert.Equal("Series2", model.Series[1].Name);
            Assert.Equal(new List<double> { 3, 4 }, model.Series[1].Values);
            Assert.Equal(ChartType.Column, model.Type);
        }

        [Fact]
        public void RecognizesPie_WhenSingleSeries()
        {
            double[,] data = {
                { double.NaN, 5, 6, 7 },
                { 0, 1, 2, 3 }
            };
            var bytes = BuildBiffMatrix(data, rowLabels: new[] { "OnlySeries" }, colLabels: new[] { "A", "B", "C" });
            var model = new ChartModel { SourceBytes = bytes, SourceStreamName = "Chart1" };
            BiffChartScanner.TryPopulateChart(model);
            Assert.Equal(ChartType.Pie, model.Type);
        }

        [Fact]
        public void PicksCategoriesFromFirstColumn_WhenOrientationSuggests()
        {
            // sheet has no header row and two data columns so we exercise
            // orientation logic and avoid the "single series = pie" heuristic.
            double[,] data = {
                { double.NaN, double.NaN, double.NaN },
                { double.NaN, 10, 11 },
                { double.NaN, 20, 21 },
                { double.NaN, 30, 31 }
            };
            var bytes = BuildBiffMatrix(data);

            // add explicit labels in column zero at rows 1..3
            using var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            using (var w = new BinaryWriter(ms, Encoding.Default, leaveOpen: true))
            {
                void WriteLabel(int row, string text)
                {
                    var b = Encoding.Default.GetBytes(text);
                    w.Write((ushort)0x0204);
                    w.Write((ushort)(8 + 1 + b.Length));
                    w.Write((ushort)row);
                    w.Write((ushort)0);
                    w.Write((ushort)0);
                    w.Write((ushort)b.Length);
                    w.Write((byte)0);
                    w.Write(b);
                }

                WriteLabel(1, "Cat1");
                WriteLabel(2, "Cat2");
                WriteLabel(3, "Cat3");
            }

            var final = ms.ToArray();
            var model2 = new ChartModel { SourceBytes = final, SourceStreamName = "Chart1" };
            BiffChartScanner.TryPopulateChart(model2);
            Assert.Equal(3, model2.Categories.Count);
            Assert.Equal("Cat1", model2.Categories[0]);
            Assert.Equal(ChartType.Column, model2.Type);
        }
    }
}
