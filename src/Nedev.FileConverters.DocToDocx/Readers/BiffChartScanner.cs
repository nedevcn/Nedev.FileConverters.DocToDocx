using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// A lightweight binary scanner that attempts to extract simple series data
/// from BIFF (Excel) streams embedded within OLE objects.
/// </summary>
public static class BiffChartScanner
{
    public static void TryPopulateChart(ChartModel model)
    {
        var bytes = model.SourceBytes;
        if (bytes == null || bytes.Length < 8) return;

        // Check if it's an OLE CFB container (D0 CF 11 E0)
        if (bytes[0] == 0xD0 && bytes[1] == 0xCF && bytes[2] == 0x11 && bytes[3] == 0xE0)
        {
            try
            {
                using var ms = new MemoryStream(bytes);
                using var cfb = new CfbReader(ms, leaveOpen: true);
                
                // Typical Excel embedded workbook stream names
                string? targetStream = cfb.StreamNames.FirstOrDefault(n => 
                    n.Equals("Workbook", StringComparison.OrdinalIgnoreCase) ||
                    n.Equals("Book", StringComparison.OrdinalIgnoreCase));
                    
                if (targetStream != null)
                {
                    var biffBytes = cfb.GetStreamBytes(targetStream);
                    ParseBiffStream(biffBytes, model);
                    return;
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to inspect embedded chart source '{model.SourceStreamName ?? "<unknown>"}' as a CFB workbook.", ex);
            }
        }
        
        // If not OLE, check if it's a raw BIFF stream (BOF record starts with 09 08, 09 04, etc.)
        if (bytes[0] == 0x09 && (bytes[1] == 0x08 || bytes[1] == 0x04 || bytes[1] == 0x02 || bytes[1] == 0x00))
        {
            ParseBiffStream(bytes, model);
        }
    }

    private static void ParseBiffStream(byte[] biffBytes, ChartModel model)
    {
        var cells = new Dictionary<(int r, int c), double>();
        var strings = new Dictionary<(int r, int c), string>();
        var sst = new List<string>();
        string? firstSheetName = null;
        
        using var ms = new MemoryStream(biffBytes);
        using var reader = new BinaryReader(ms);
        
        try
        {
            while (ms.Position + 4 <= ms.Length)
            {
                ushort id = reader.ReadUInt16();
                ushort length = reader.ReadUInt16();

                if (ms.Position + length > ms.Length) break;
                
                long nextPos = ms.Position + length;

                if (id == 0x00FC) // SST
                {
                    ParseSst(reader, length, sst);
                }
                else if (id == 0x0085) // BOUNDSHEET
                {
                    if (string.IsNullOrEmpty(firstSheetName))
                    {
                        firstSheetName = ParseBoundSheetName(reader, length);
                    }
                }
                else if (id == 0x00FD) // LABELSST
                {
                    if (length >= 10)
                    {
                        int row = reader.ReadUInt16();
                        int col = reader.ReadUInt16();
                        reader.ReadUInt16(); // xf
                        uint sstIndex = reader.ReadUInt32();
                        if (sstIndex < sst.Count)
                            strings[(row, col)] = sst[(int)sstIndex];
                    }
                }
                else if (id == 0x0203) // NUMBER
                {
                    if (length >= 14)
                    {
                        int row = reader.ReadUInt16();
                        int col = reader.ReadUInt16();
                        reader.ReadUInt16(); // xf
                        double val = reader.ReadDouble();
                        cells[(row, col)] = val;
                    }
                }
                else if (id == 0x027E) // RK
                {
                    if (length >= 10)
                    {
                        int row = reader.ReadUInt16();
                        int col = reader.ReadUInt16();
                        reader.ReadUInt16(); // xf
                        uint rk = reader.ReadUInt32();
                        cells[(row, col)] = DecodeRk(rk);
                    }
                }
                else if (id == 0x00BD) // MULRK
                {
                    ParseMulRk(reader, length, cells);
                }
                else if (id == 0x0204) // LABEL
                {
                    if (length >= 8)
                    {
                        int row = reader.ReadUInt16();
                        int col = reader.ReadUInt16();
                        reader.ReadUInt16(); // xf
                        ushort strLen = reader.ReadUInt16();
                        // Extract string - simplified
                        if (length > 8)
                        {
                            byte flag = reader.ReadByte();
                            bool isUnicode = (flag & 0x01) != 0;
                            if (isUnicode)
                                strings[(row, col)] = Encoding.Unicode.GetString(reader.ReadBytes(Math.Min(strLen * 2, length - 9)));
                            else
                                strings[(row, col)] = Encoding.Default.GetString(reader.ReadBytes(Math.Min((int)strLen, length - 9)));
                        }
                    }
                }
                else if (id == 0x0205) // BOOLERR
                {
                    if (length >= 8)
                    {
                        int row = reader.ReadUInt16();
                        int col = reader.ReadUInt16();
                        reader.ReadUInt16(); // xf
                        byte val = reader.ReadByte();
                        byte isError = reader.ReadByte();
                        if (isError == 0)
                        {
                            cells[(row, col)] = val != 0 ? 1d : 0d;
                        }
                    }
                }
                
                ms.Position = nextPos;
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Stopped parsing BIFF chart stream '{model.SourceStreamName ?? "<unknown>"}' due to malformed or truncated record data.", ex);
        }
        
        if (cells.Count == 0) return;
        
        int maxRow = cells.Keys.Max(k => k.r);
        int maxCol = cells.Keys.Max(k => k.c);
        
        // Decide whether categories live in the first row (default) or first column.
        bool useRowCategories = true;
        if (maxRow > maxCol)
        {
            bool hasRowCat = Enumerable.Range(1, maxCol)
                .Any(c => strings.ContainsKey((0, c)) || cells.ContainsKey((0, c)));
            bool hasColCat = Enumerable.Range(1, maxRow)
                .Any(r => strings.ContainsKey((r, 0)) || cells.ContainsKey((r, 0)));
            if (!hasRowCat && hasColCat)
                useRowCategories = false;
        }
        
        List<string> categories = new();
        var seriesList = new List<ChartSeries>();

        if (useRowCategories)
        {
            for (int c = 1; c <= maxCol; c++)
            {
                if (strings.TryGetValue((0, c), out var catName))
                    categories.Add(catName);
                else if (cells.TryGetValue((0, c), out var catVal))
                    categories.Add(catVal.ToString("G"));
                else
                    categories.Add($"Category {c}");
            }

            for (int r = 1; r <= maxRow; r++)
            {
                var values = new List<double>();
                bool hasValue = false;
                for (int c = 1; c <= maxCol; c++)
                {
                    if (cells.TryGetValue((r, c), out var v))
                    {
                        values.Add(v);
                        hasValue = true;
                    }
                    else
                    {
                        values.Add(0);
                    }
                }

                if (hasValue)
                {
                    string seriesName = $"Series {r}";
                    if (strings.TryGetValue((r, 0), out var sName)) seriesName = sName;
                    else if (cells.TryGetValue((r, 0), out var sVal)) seriesName = sVal.ToString("G");

                    seriesList.Add(new ChartSeries { Name = seriesName, Values = values });
                }
            }
        }
        else
        {
            // categories come from first column, series run across columns
            for (int r = 1; r <= maxRow; r++)
            {
                if (strings.TryGetValue((r, 0), out var catName))
                    categories.Add(catName);
                else if (cells.TryGetValue((r, 0), out var catVal))
                    categories.Add(catVal.ToString("G"));
                else
                    categories.Add($"Category {r}");
            }

            for (int c = 1; c <= maxCol; c++)
            {
                var values = new List<double>();
                bool hasValue = false;
                for (int r = 1; r <= maxRow; r++)
                {
                    if (cells.TryGetValue((r, c), out var v))
                    {
                        values.Add(v);
                        hasValue = true;
                    }
                    else
                    {
                        values.Add(0);
                    }
                }

                if (hasValue)
                {
                    string seriesName = $"Series {c}";
                    if (strings.TryGetValue((0, c), out var sName)) seriesName = sName;
                    else if (cells.TryGetValue((0, c), out var sVal)) seriesName = sVal.ToString("G");

                    seriesList.Add(new ChartSeries { Name = seriesName, Values = values });
                }
            }
        }

        if (seriesList.Count > 0)
        {
            model.Categories = categories;
            model.Series = seriesList;
            if (seriesList.Count <= 1)
            {
                model.ShowLegend = false;
            }

            if (strings.TryGetValue((0, 0), out var topLeftText) && !string.IsNullOrWhiteSpace(topLeftText) && string.IsNullOrWhiteSpace(model.CategoryAxisTitle))
            {
                model.CategoryAxisTitle = topLeftText;
            }

            if ((string.IsNullOrWhiteSpace(model.Title) || string.Equals(model.Title, model.SourceStreamName, StringComparison.OrdinalIgnoreCase))
                && !string.IsNullOrWhiteSpace(firstSheetName))
            {
                model.Title = firstSheetName;
            }

            // heuristic type detection
            if (seriesList.Count == 1 && categories.Count > 1)
            {
                model.Type = ChartType.Pie;
            }
            else if (model.SourceStreamName != null)
            {
                var name = model.SourceStreamName.ToLowerInvariant();
                if (name.Contains("pie")) model.Type = ChartType.Pie;
                else if (name.Contains("line")) model.Type = ChartType.Line;
                else if (name.Contains("bar")) model.Type = ChartType.Bar;
                else if (name.Contains("area")) model.Type = ChartType.Area;
                else if (name.Contains("scatter")) model.Type = ChartType.Scatter;
            }
        }
    }

    private static void ParseMulRk(BinaryReader reader, int length, Dictionary<(int r, int c), double> cells)
    {
        if (length < 6)
            return;

        int row = reader.ReadUInt16();
        int firstCol = reader.ReadUInt16();
        int rkEntryCount = (length - 6) / 6;
        for (int i = 0; i < rkEntryCount; i++)
        {
            reader.ReadUInt16(); // xf
            uint rk = reader.ReadUInt32();
            cells[(row, firstCol + i)] = DecodeRk(rk);
        }

        reader.ReadUInt16(); // last col
    }

    private static string? ParseBoundSheetName(BinaryReader reader, int length)
    {
        if (length < 8)
            return null;

        reader.ReadUInt32();
        reader.ReadByte();
        reader.ReadByte();
        byte nameLength = reader.ReadByte();
        byte flags = reader.ReadByte();
        bool isUnicode = (flags & 0x01) != 0;

        if (nameLength == 0)
            return null;

        return isUnicode
            ? Encoding.Unicode.GetString(reader.ReadBytes(nameLength * 2))
            : Encoding.Default.GetString(reader.ReadBytes(nameLength));
    }

    private static void ParseSst(BinaryReader reader, int length, List<string> sst)
    {
        try
        {
            uint totalUnique = reader.ReadUInt32();
            for (int i = 0; i < totalUnique; i++)
            {
                if (reader.BaseStream.Position >= reader.BaseStream.Length) break;
                ushort charLen = reader.ReadUInt16();
                byte flag = reader.ReadByte();
                bool isUnicode = (flag & 0x01) != 0;
                
                string s;
                if (isUnicode)
                    s = Encoding.Unicode.GetString(reader.ReadBytes(charLen * 2));
                else
                    s = Encoding.Default.GetString(reader.ReadBytes(charLen));
                
                sst.Add(s);
                
                // Skip formatting runs and phonetic data if any
                bool hasFormattingRuns = (flag & 0x08) != 0;
                bool hasPhoneticData = (flag & 0x04) != 0;
                if (hasFormattingRuns) reader.ReadBytes(reader.ReadUInt16() * 4);
                if (hasPhoneticData) reader.ReadBytes((int)reader.ReadUInt32());
            }
        }
        catch (Exception ex)
        {
            Logger.Warning("Failed to parse BIFF shared string table completely; continuing with partial chart labels.", ex);
        }
    }

    /// <summary>
    /// Decodes an Excel RK compressed number format.
    /// </summary>
    private static double DecodeRk(uint rk)
    {
        bool isFloat = (rk & 0x02) == 0;
        bool isDiv100 = (rk & 0x01) != 0;
        double value;

        if (isFloat)
        {
            ulong bits = ((ulong)(rk & 0xFFFFFFFC)) << 32;
            value = BitConverter.Int64BitsToDouble((long)bits);
        }
        else
        {
            int intVal = (int)(rk >> 2);
            value = intVal;
        }

        if (isDiv100)
            value /= 100.0;

        return value;
    }
}
