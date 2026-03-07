using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Writers;

public class NumberingWriter
{
    private readonly XmlWriter _writer;
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public NumberingWriter(XmlWriter writer)
    {
        _writer = writer;
    }

    public void WriteNumbering(DocumentModel document)
    {
        _writer.WriteStartDocument();
        _writer.WriteStartElement("w", "numbering", WNs);
        _writer.WriteAttributeString("xmlns", "w", null, WNs);
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        var numberingDefinitions = BuildNumberingDefinitions(document);

        if (numberingDefinitions.Count > 0)
        {
            foreach (var numDef in numberingDefinitions)
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

    private static List<NumberingDefinition> BuildNumberingDefinitions(DocumentModel document)
    {
        var definitions = document.NumberingDefinitions
            .Where(definition => definition.Id > 0)
            .GroupBy(definition => definition.Id)
            .Select(group => group.First())
            .OrderBy(definition => definition.Id)
            .ToList();

        var usedListIds = document.Paragraphs
            .Select(paragraph => paragraph.Properties?.ListFormatId ?? paragraph.ListFormatId)
            .Where(listId => listId > 0)
            .Distinct()
            .OrderBy(listId => listId)
            .ToList();

        foreach (var listId in usedListIds)
        {
            if (definitions.Any(definition => definition.Id == listId))
            {
                continue;
            }

            definitions.Add(CreateFallbackDefinition(listId));
        }

        definitions.Sort((left, right) => left.Id.CompareTo(right.Id));
        return definitions;
    }

    private static NumberingDefinition CreateFallbackDefinition(int listId)
    {
        var definition = new NumberingDefinition
        {
            Id = listId
        };

        for (int level = 0; level < 9; level++)
        {
            definition.Levels.Add(new NumberingLevel
            {
                Level = level,
                NumberFormat = NumberFormat.Decimal,
                Text = $"%{level + 1}.",
                Start = 1
            });
        }

        return definition;
    }

    private void WriteDefaultNumbering()
    {
        _writer.WriteStartElement("w", "abstractNum", WNs);
        _writer.WriteAttributeString("w", "abstractNumId", WNs, "0");

        _writer.WriteStartElement("w", "nsid", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "00000000");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "multiLevelType", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "hybridMultilevel");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "tmpl", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "00000000");
        _writer.WriteEndElement();

        for (int lvl = 0; lvl < 9; lvl++)
        {
            WriteLevel(lvl, NumberFormat.Decimal, $"%{lvl + 1}.", 1);
        }

        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "num", WNs);
        _writer.WriteAttributeString("w", "numId", WNs, "1");
        _writer.WriteStartElement("w", "abstractNumId", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "0");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
    }

    private void WriteAbstractNum(NumberingDefinition numDef, int abstractNumId)
    {
        _writer.WriteStartElement("w", "abstractNum", WNs);
        _writer.WriteAttributeString("w", "abstractNumId", WNs, abstractNumId.ToString());

        _writer.WriteStartElement("w", "nsid", WNs);
        _writer.WriteAttributeString("w", "val", WNs, Convert.ToString(numDef.Id, 16).PadLeft(8, '0').ToUpper());
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "multiLevelType", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "hybridMultilevel");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "tmpl", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "00000000");
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
        _writer.WriteStartElement("w", "lvl", WNs);
        _writer.WriteAttributeString("w", "ilvl", WNs, level.ToString());

        _writer.WriteStartElement("w", "start", WNs);
        _writer.WriteAttributeString("w", "val", WNs, start.ToString());
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "numFmt", WNs);
        _writer.WriteAttributeString("w", "val", WNs, GetNumberFormatValue(numFmt));
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlText", WNs);
        _writer.WriteAttributeString("w", "val", WNs, prefix);
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlJc", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "left");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "pPr", WNs);
        _writer.WriteStartElement("w", "ind", WNs);
        _writer.WriteAttributeString("w", "left", WNs, (720 + level * 720).ToString());
        _writer.WriteAttributeString("w", "hanging", WNs, "360");
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // w:pPr

        _writer.WriteEndElement(); // w:lvl
    }

    private void WriteNum(int numId, int abstractNumId)
    {
        _writer.WriteStartElement("w", "num", WNs);
        _writer.WriteAttributeString("w", "numId", WNs, numId.ToString());

        _writer.WriteStartElement("w", "abstractNumId", WNs);
        _writer.WriteAttributeString("w", "val", WNs, abstractNumId.ToString());
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
