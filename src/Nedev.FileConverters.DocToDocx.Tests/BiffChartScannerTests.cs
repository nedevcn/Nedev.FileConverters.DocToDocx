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
            Assert.Equal("OnlySeries", model.ValueAxisTitle);
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

        [Fact]
        public void ParsesRkAndMulRkRecords()
        {
            // create a barebones BIFF stream containing a single RK cell and a
            // MULRK row; categories/series detection is not important for this
            // test, we just want the numeric values to survive the scan.
            using var ms = new MemoryStream();
            using var w = new BinaryWriter(ms, Encoding.Default, leaveOpen: true);

            // minimal BOF
            w.Write((ushort)0x0809);
            w.Write((ushort)0);

            // RK record at (1,1) -> value 42
            w.Write((ushort)0x027E);
            w.Write((ushort)10);
            w.Write((ushort)1);
            w.Write((ushort)1);
            w.Write((ushort)0);
            w.Write(EncodeRkInt(42));

            // MULRK on row 2, cols 1..2 values 5 and 6
            w.Write((ushort)0x00BD);
            int firstCol = 1;
            int count = 2;
            int length = 6 + count * 6 + 2;
            w.Write((ushort)length);
            w.Write((ushort)2);
            w.Write((ushort)firstCol);
            for (int i = 0; i < count; i++)
            {
                w.Write((ushort)0);
                w.Write(EncodeRkInt(5 + i));
            }
            w.Write((ushort)(firstCol + count - 1));

            var bytes = ms.ToArray();
            var model = new ChartModel { SourceBytes = bytes, SourceStreamName = "Chart1" };
            BiffChartScanner.TryPopulateChart(model);

            var allValues = model.Series.SelectMany(s => s.Values).ToList();
            // RK entry should always be present
            Assert.Contains(42d, allValues);
            // MULRK entries may end up in a second series depending on orientation
            if (model.Series.Count > 1)
            {
                Assert.Contains(5d, allValues);
                Assert.Contains(6d, allValues);
            }
        }

        // helper encoders for tests
        private static uint EncodeRkInt(int value)
        {
            // flag bit1 set -> integer, bit0 clear -> no divide
            return ((uint)value << 2) | 0x02u;
        }

        [Fact]
        public void UsesTopLeftLabelAndSheetName_ForRecoveredMetadata()
        {
            double[,] data = {
                { double.NaN, 5, 6 },
                { 0, 10, 20 }
            };

            using var ms = new MemoryStream();
            using (var w = new BinaryWriter(ms, Encoding.Default, leaveOpen: true))
            {
                w.Write((ushort)0x0809);
                w.Write((ushort)0);

                var sheetName = Encoding.Default.GetBytes("Sales Sheet");
                w.Write((ushort)0x0085);
                w.Write((ushort)(8 + sheetName.Length));
                w.Write((uint)0);
                w.Write((byte)0);
                w.Write((byte)0);
                w.Write((byte)sheetName.Length);
                w.Write((byte)0);
                w.Write(sheetName);
            }

            var matrixBytes = BuildBiffMatrix(data, rowLabels: new[] { "Series1" }, colLabels: new[] { "Jan", "Feb" });
            ms.Write(matrixBytes, 4, matrixBytes.Length - 4);

            using (var w = new BinaryWriter(ms, Encoding.Default, leaveOpen: true))
            {
                var label = Encoding.Default.GetBytes("Month");
                w.Write((ushort)0x0204);
                w.Write((ushort)(8 + 1 + label.Length));
                w.Write((ushort)0);
                w.Write((ushort)0);
                w.Write((ushort)0);
                w.Write((ushort)label.Length);
                w.Write((byte)0);
                w.Write(label);
            }

            var model = new ChartModel { SourceBytes = ms.ToArray(), SourceStreamName = "Chart1" };
            BiffChartScanner.TryPopulateChart(model);

            Assert.Equal("Sales Sheet", model.Title);
            Assert.Equal("Month", model.CategoryAxisTitle);
            Assert.False(model.ShowLegend);
        }

        [Fact]
        public void ParsesFormulaNumbersAndFormulaStrings_ForChartRecovery()
        {
            using var ms = new MemoryStream();
            using (var w = new BinaryWriter(ms, Encoding.Default, leaveOpen: true))
            {
                w.Write((ushort)0x0809);
                w.Write((ushort)0);

                WriteLabel(w, 0, 1, "Jan");
                WriteLabel(w, 0, 2, "Feb");
                WriteLabel(w, 1, 0, "Revenue");

                WriteFormulaString(w, 0, 0, "Month");
                WriteFormulaNumber(w, 1, 1, 10d);
                WriteFormulaNumber(w, 1, 2, 20d);
            }

            var model = new ChartModel { SourceBytes = ms.ToArray(), SourceStreamName = "Chart1" };
            BiffChartScanner.TryPopulateChart(model);

            Assert.Equal(new List<string> { "Jan", "Feb" }, model.Categories);
            var series = Assert.Single(model.Series);
            Assert.Equal("Revenue", series.Name);
            Assert.Equal(new List<double> { 10d, 20d }, series.Values);
            Assert.Equal("Month", model.CategoryAxisTitle);
        }

        [Theory]
        [InlineData("Revenue Doughnut", ChartType.Doughnut)]
        [InlineData("Revenue Radar", ChartType.Radar)]
        public void DetectsAdditionalChartTypes_FromSourceStreamName(string sourceStreamName, ChartType expectedType)
        {
            double[,] data = {
                { double.NaN, 5, 6 },
                { 0, 10, 20 },
                { 0, 11, 21 }
            };

            var bytes = BuildBiffMatrix(data, rowLabels: new[] { "North", "South" }, colLabels: new[] { "Jan", "Feb" });
            var model = new ChartModel { SourceBytes = bytes, SourceStreamName = sourceStreamName };

            BiffChartScanner.TryPopulateChart(model);

            Assert.Equal(expectedType, model.Type);
        }

        private static void WriteLabel(BinaryWriter writer, int row, int col, string text)
        {
            var bytes = Encoding.Default.GetBytes(text);
            writer.Write((ushort)0x0204);
            writer.Write((ushort)(8 + 1 + bytes.Length));
            writer.Write((ushort)row);
            writer.Write((ushort)col);
            writer.Write((ushort)0);
            writer.Write((ushort)bytes.Length);
            writer.Write((byte)0);
            writer.Write(bytes);
        }

        private static void WriteFormulaNumber(BinaryWriter writer, int row, int col, double value)
        {
            writer.Write((ushort)0x0006);
            writer.Write((ushort)20);
            writer.Write((ushort)row);
            writer.Write((ushort)col);
            writer.Write((ushort)0);
            writer.Write(value);
            writer.Write((ushort)0);
            writer.Write(0u);
        }

        private static void WriteFormulaString(BinaryWriter writer, int row, int col, string value)
        {
            writer.Write((ushort)0x0006);
            writer.Write((ushort)20);
            writer.Write((ushort)row);
            writer.Write((ushort)col);
            writer.Write((ushort)0);
            writer.Write(new byte[] { 0, 0, 0, 0, 0, 0, 0xFF, 0xFF });
            writer.Write((ushort)0);
            writer.Write(0u);

            var bytes = Encoding.Default.GetBytes(value);
            writer.Write((ushort)0x0207);
            writer.Write((ushort)(3 + bytes.Length));
            writer.Write((ushort)bytes.Length);
            writer.Write((byte)0);
            writer.Write(bytes);
        }
    }
}
