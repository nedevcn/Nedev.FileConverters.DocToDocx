using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Collections.Generic;

namespace Nedev.DocToDocx.Readers;

/// <summary>
/// Parser for MTEF (MathType Equation Format) binary data.
/// Converts legacy Equation Editor 3.0 data into OOXML Math (OMML).
/// Reference: [MS-MTMTEF] or Equation Editor documentation.
/// </summary>
public class MtefReader
{
    private readonly byte[] _data;
    private int _pos;

    public MtefReader(byte[] data)
    {
        _data = data;
        _pos = 0;
    }

    /// <summary>
    /// Parses MTEF data and returns an OMML fragment as an XML string.
    /// </summary>
    public string? ConvertToOmml()
    {
        if (_data == null || _data.Length < 5) return null;

        try
        {
            // OLE storage for Equation Native usually starts with an OLE header
            // followed by MTEF data. Check for MTEF version.
            // Search for MTEF header (0x03 0x01 0x01 0x03)
            int startPos = -1;
            for (int i = 0; i < _data.Length - 4; i++)
            {
                if (_data[i] == 0x03 && _data[i + 1] == 0x01 && _data[i + 2] == 0x01 && _data[i + 3] == 0x03)
                {
                    startPos = i;
                    break;
                }
            }
            if (startPos == -1) startPos = 0;
            _pos = startPos;

            // Skip MTEF header (already checked)
            _pos += 5; // Skip v3 header

            var sb = new StringBuilder();
            using (var writer = XmlWriter.Create(sb, new XmlWriterSettings { OmitXmlDeclaration = true, ConformanceLevel = ConformanceLevel.Fragment }))
            {
                writer.WriteStartElement("m", "oMath", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement();
            }

            return sb.ToString();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error parsing MTEF: {ex.Message}");
            return null;
        }
    }

    private void ParseRecords(XmlWriter writer)
    {
        while (_pos < _data.Length)
        {
            byte tag = _data[_pos++];
            if (tag == 0) break; // END tag

            // Bits 0-3: Tag type, Bits 4-7: Flags
            byte cmd = (byte)(tag & 0x0F);
            byte options = (byte)(tag >> 4);

            switch (cmd)
            {
                case 1: // LINE
                    ParseLine(writer, options);
                    break;
                case 2: // CHAR
                    ParseChar(writer, options);
                    break;
                case 3: // TMPL (Template like Fractions, Radicals)
                    ParseTemplate(writer, options);
                    break;
                case 4: // PILE
                    ParsePile(writer, options);
                    break;
                case 5: // MATRIX
                    SkipRecord(); 
                    break;
                case 6: // EMBELL (Overbar, etc.)
                    SkipRecord();
                    break;
                default:
                    // Unknown tag, stop parsing to avoid corruption
                    return;
            }
        }
    }

    private void ParseLine(XmlWriter writer, byte options)
    {
        // Skip line options (val, spacing)
        if ((options & 0x01) != 0) _pos++; // halign
        if ((options & 0x02) != 0) _pos++; // valign
        
        ParseRecords(writer);
    }

    private void ParsePile(XmlWriter writer, byte options)
    {
        _pos++; // halign
        _pos++; // valign
        ParseRecords(writer);
    }

    private void ParseChar(XmlWriter writer, byte options)
    {
        // Tag + [variation] + [font] + [char]
        if ((options & 0x01) != 0) _pos++; // typeface
        if ((options & 0x02) != 0) _pos++; // char size
        
        // MTEF characters are usually 16-bit
        if (_pos + 2 > _data.Length) return;
        short chValue = BitConverter.ToInt16(_data, _pos);
        _pos += 2;

        char c = (char)chValue;
        
        writer.WriteStartElement("m", "r", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        writer.WriteStartElement("m", "t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        writer.WriteString(MapChar(c).ToString());
        writer.WriteEndElement(); // m:t
        writer.WriteEndElement(); // m:r
    }

    private void ParseTemplate(XmlWriter writer, byte options)
    {
        byte type = _data[_pos++]; // Template type
        _pos++; // variation
        _pos++; // options

        switch (type)
        {
            case 0: // Fraction
                writer.WriteStartElement("m", "f", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                
                writer.WriteStartElement("m", "num", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:num

                writer.WriteStartElement("m", "den", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:den

                writer.WriteEndElement(); // m:f
                break;

            case 3: // Radical
                writer.WriteStartElement("m", "rad", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                writer.WriteStartElement("m", "radPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                writer.WriteStartElement("m", "degHide", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                writer.WriteEndElement();
                writer.WriteEndElement();
                
                writer.WriteStartElement("m", "deg", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                writer.WriteEndElement();

                writer.WriteStartElement("m", "e", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:e

                writer.WriteEndElement(); // m:rad
                break;

            case 6: // Subscript
                writer.WriteStartElement("m", "sSub", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                
                writer.WriteStartElement("m", "e", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:e

                writer.WriteStartElement("m", "sub", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:sub

                writer.WriteEndElement(); // m:sSub
                break;

            case 7: // Superscript
                writer.WriteStartElement("m", "sSup", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                
                writer.WriteStartElement("m", "e", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:e

                writer.WriteStartElement("m", "sup", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:sup

                writer.WriteEndElement(); // m:sSup
                break;

            case 8: // Sub/Superscript
                writer.WriteStartElement("m", "sSubSup", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                
                writer.WriteStartElement("m", "e", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:e

                writer.WriteStartElement("m", "sub", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:sub

                writer.WriteStartElement("m", "sup", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                ParseRecords(writer);
                writer.WriteEndElement(); // m:sup

                writer.WriteEndElement(); // m:sSubSup
                break;

            default:
                // Unsupported template, just parse its content as a sequence
                ParseRecords(writer);
                break;
        }
    }

    private void SkipRecord()
    {
        // Basic skip logic: records end with END (0)
        int depth = 1;
        while (_pos < _data.Length && depth > 0)
        {
            byte tag = _data[_pos++];
            if (tag == 0) depth--;
            else
            {
                byte cmd = (byte)(tag & 0x0F);
                // Tags that start a new scope
                if (cmd == 1 || cmd == 3 || cmd == 4 || cmd == 5) depth++;
                // Skip payload for characters
                if (cmd == 2)
                {
                    byte options = (byte)(tag >> 4);
                    if ((options & 0x01) != 0) _pos++;
                    if ((options & 0x02) != 0) _pos++;
                    _pos += 2;
                }
            }
        }
    }

    private char MapChar(char c)
    {
        // Simple mapping for common symbols if needed
        return c;
    }
}
