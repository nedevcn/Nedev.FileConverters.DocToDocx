using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Readers;

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
            catch { /* Ignore CFB parse errors from partial streams */ }
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

                if (id == 0x0203) // NUMBER
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
                else if (id == 0x0204) // LABEL
                {
                    if (length >= 8)
                    {
                        int row = reader.ReadUInt16();
                        int col = reader.ReadUInt16();
                        reader.ReadUInt16(); // xf
                        ushort strLen = reader.ReadUInt16();
                        int readLen = Math.Min((int)strLen, length - 8);
                        if (readLen > 0)
                        {
                            byte[] strBytes = reader.ReadBytes(readLen);
                            strings[(row, col)] = Encoding.Default.GetString(strBytes);
                        }
                    }
                }
                // (Note: ignoring LABELSST / SST for extreme lightweight parsing;
                // chart labels will fallback to default "Category N" if strings are absent)
                
                ms.Position = nextPos;
            }
        }
        catch { /* Stop parsing if stream is abruptly truncated */ }
        
        if (cells.Count == 0) return;
        
        int maxRow = cells.Keys.Max(k => k.r);
        int maxCol = cells.Keys.Max(k => k.c);
        
        // Assume Row 0 holds category names, Col 0 holds series names. Data is in (1..maxRow, 1..maxCol).
        var categories = new List<string>();
        for (int c = 1; c <= maxCol; c++)
        {
            if (strings.TryGetValue((0, c), out var catName))
                categories.Add(catName);
            else if (cells.TryGetValue((0, c), out var catVal))
                categories.Add(catVal.ToString("G"));
            else
                categories.Add($"Category {c}");
        }
        
        var seriesList = new List<ChartSeries>();
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
        
        if (seriesList.Count > 0)
        {
            model.Categories = categories;
            model.Series = seriesList;
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
