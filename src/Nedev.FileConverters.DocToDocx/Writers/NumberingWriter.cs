using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Writers;

public class NumberingWriter
{
    private readonly XmlWriter _writer;

    public NumberingWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    public void WriteNumbering(DocumentModel document)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "numbering", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        if (document.NumberingDefinitions.Count > 0)
        {
            // Use NumberingDefinition.Id consistently as both abstractNumId and numId
            foreach (var numDef in document.NumberingDefinitions)
            {
                var id = numDef.Id;
                if (id <= 0)
                {
                    continue;
                }

                WriteAbstractNum(numDef, id);
                WriteNum(id, id);
            }
        }
        else
        {
            WriteDefaultNumbering();
        }

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }

    private void WriteDefaultNumbering()
    {
        _writer.WriteStartElement("w", "abstractNum");
        _writer.WriteAttributeString("w", "abstractNumId", null, "0");

        _writer.WriteStartElement("w", "nsid");
        _writer.WriteAttributeString("w", "val", null, "00000000");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "multiLevelType");
        _writer.WriteAttributeString("w", "val", null, "hybridMultilevel");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "tmpl");
        _writer.WriteAttributeString("w", "val", null, "00000000");
        _writer.WriteEndElement();

        for (int lvl = 0; lvl < 9; lvl++)
        {
            WriteLevel(lvl, NumberFormat.Decimal, " ", lvl + 1);
        }

        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "num");
        _writer.WriteAttributeString("w", "numId", null, "1");
        _writer.WriteStartElement("w", "abstractNumId");
        _writer.WriteAttributeString("w", "val", null, "0");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }

    private void WriteAbstractNum(NumberingDefinition numDef, int abstractNumId)
    {
        _writer.WriteStartElement("w", "abstractNum");
        _writer.WriteAttributeString("w", "abstractNumId", null, abstractNumId.ToString());

        _writer.WriteStartElement("w", "nsid");
        _writer.WriteAttributeString("w", "val", Convert.ToString(numDef.Id, 16).PadLeft(8, '0').ToUpper());
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "multiLevelType");
        _writer.WriteAttributeString("w", "val", null, "hybridMultilevel");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "tmpl");
        _writer.WriteAttributeString("w", "val", null, "00000000");
        _writer.WriteEndElement();

        // Write levels from actual definition data
        if (numDef.Levels.Count > 0)
        {
            foreach (var level in numDef.Levels)
            {
                WriteLevel(level.Level, level.NumberFormat, level.Text ?? $"%{level.Level + 1}.", level.Start);
            }
        }
        else
        {
            // Fallback: write a single bullet level
            WriteLevel(0, NumberFormat.Bullet, "\u00B7", 1);
        }

        _writer.WriteEndElement(); // w:abstractNum
    }

    private void WriteLevel(int level, NumberFormat numFmt, string prefix, int start)
    {
        _writer.WriteStartElement("w", "lvl");
        _writer.WriteAttributeString("w", "ilvl", null, level.ToString());

        _writer.WriteStartElement("w", "start");
        _writer.WriteAttributeString("w", "val", null, start.ToString());
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "numFmt");
        _writer.WriteAttributeString("w", "val", null, GetNumberFormatValue(numFmt));
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlText");
        _writer.WriteAttributeString("w", "val", null, prefix);
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlJc");
        _writer.WriteAttributeString("w", "val", null, "left");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "pPr");
        _writer.WriteStartElement("w", "ind");
        _writer.WriteAttributeString("w", "left", null, (720 + level * 720).ToString());
        _writer.WriteAttributeString("w", "hanging", null, "360");
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // w:pPr

        _writer.WriteEndElement(); // w:lvl
    }

    private void WriteNum(int numId, int abstractNumId)
    {
        _writer.WriteStartElement("w", "num");
        _writer.WriteAttributeString("w", "numId", numId.ToString());

        _writer.WriteStartElement("w", "abstractNumId");
        _writer.WriteAttributeString("w", "val", abstractNumId.ToString());
        _writer.WriteEndElement();

        _writer.WriteEndElement();
    }

    private string GetNumberFormatValue(NumberFormat format)
    {
        return format switch
        {
            NumberFormat.Bullet => "bullet",
            NumberFormat.Decimal => "decimal",
            NumberFormat.LowerRoman => "lowerRoman",
            NumberFormat.UpperRoman => "upperRoman",
            NumberFormat.LowerLetter => "lowerLetter",
            NumberFormat.UpperLetter => "upperLetter",
            NumberFormat.Ordinal => "ordinal",
            NumberFormat.CardinalText => "cardinalText",
            NumberFormat.OrdinalText => "ordinalText",
            NumberFormat.Hex => "hex",
            NumberFormat.Chicago => "chicago",
            NumberFormat.IdeographDigital => "ideographDigital",
            NumberFormat.JapaneseCounting => "japaneseCounting",
            NumberFormat.Aiueo => "aiueo",
            NumberFormat.Iroha => "iroha",
            NumberFormat.DecimalFullWidth => "decimalFullWidth",
            NumberFormat.DecimalHalfWidth => "decimalHalfWidth",
            NumberFormat.JapaneseLegal => "japaneseLegal",
            NumberFormat.JapaneseDigitalTenThousand => "japaneseDigitalTenThousand",
            NumberFormat.DecimalEnclosedCircle => "decimalEnclosedCircle",
            NumberFormat.AiueoFullWidth => "aiueoFullWidth",
            NumberFormat.IrohaFullWidth => "irohaFullWidth",
            NumberFormat.DecimalZero => "decimalZero",
            NumberFormat.Ganada => "ganada",
            NumberFormat.Chosung => "chosung",
            NumberFormat.DecimalEnclosedFullstop => "decimalEnclosedFullstop",
            NumberFormat.DecimalEnclosedParen => "decimalEnclosedParen",
            NumberFormat.DecimalEnclosedCircleChinese => "decimalEnclosedCircleChinese",
            NumberFormat.IdeographEnclosedCircle => "ideographEnclosedCircle",
            NumberFormat.IdeographTraditional => "ideographTraditional",
            NumberFormat.IdeographZodiac => "ideographZodiac",
            NumberFormat.IdeographZodiacTraditional => "ideographZodiacTraditional",
            NumberFormat.TaiwaneseCounting => "taiwaneseCounting",
            NumberFormat.IdeographLegalTraditional => "ideographLegalTraditional",
            NumberFormat.TaiwaneseCountingThousand => "taiwaneseCountingThousand",
            NumberFormat.TaiwaneseDigital => "taiwaneseDigital",
            NumberFormat.ChineseCounting => "chineseCounting",
            NumberFormat.ChineseLegalSimplified => "chineseLegalSimplified",
            NumberFormat.ChineseLegalTraditional => "chineseLegalTraditional",
            NumberFormat.JapaneseCounting2 => "japaneseCounting2",
            NumberFormat.JapaneseDigitalHundredCount => "japaneseDigitalHundredCount",
            NumberFormat.JapaneseDigitalThousandCount => "japaneseDigitalThousandCount",
            _ => "decimal"
        };
    }
}
