using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Writers;

public class NumberingWriter
{
    private readonly XmlWriter _writer;
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private sealed class NumberingInstanceDefinition
    {
        public int NumId { get; init; }
        public int AbstractNumId { get; init; }
        public List<ListLevelOverride> LevelOverrides { get; init; } = new();
    }

    private sealed class NumberingPackage
    {
        public List<NumberingDefinition> AbstractDefinitions { get; init; } = new();
        public List<NumberingInstanceDefinition> Instances { get; init; } = new();
    }

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

        var numberingPackage = BuildNumberingDefinitions(document);

        if (numberingPackage.AbstractDefinitions.Count > 0)
        {
            foreach (var numDef in numberingPackage.AbstractDefinitions)
            {
                var id = numDef.Id;
                if (id <= 0)
                {
                    continue;
                }

                WriteAbstractNum(numDef, id);
            }

            foreach (var instance in numberingPackage.Instances)
            {
                WriteNum(instance);
            }
        }
        else
        {
            WriteDefaultNumbering();
        }

        _writer.WriteEndElement();
        _writer.WriteEndDocument();
    }

    private static NumberingPackage BuildNumberingDefinitions(DocumentModel document)
    {
        var definitions = document.NumberingDefinitions
            .Where(definition => definition.Id > 0)
            .GroupBy(definition => definition.Id)
            .Select(group => group.First())
            .ToDictionary(definition => definition.Id);

        var usedListIds = document.Paragraphs
            .Select(paragraph => (paragraph.Properties?.ListFormatId ?? 0) > 0
                ? paragraph.Properties!.ListFormatId
                : paragraph.ListFormatId)
            .Where(listId => listId > 0)
            .Distinct()
            .OrderBy(listId => listId)
            .ToList();

        var overrides = document.ListFormatOverrides
            .Where(overrideDefinition => overrideDefinition.OverrideId > 0)
            .GroupBy(overrideDefinition => overrideDefinition.OverrideId)
            .Select(group => group.First())
            .ToDictionary(overrideDefinition => overrideDefinition.OverrideId);

        foreach (var listOverride in overrides.Values)
        {
            var abstractNumId = listOverride.ListId > 0 ? listOverride.ListId : listOverride.OverrideId;
            if (!definitions.ContainsKey(abstractNumId))
            {
                definitions[abstractNumId] = CreateFallbackDefinition(abstractNumId);
            }
        }

        foreach (var listId in usedListIds)
        {
            var abstractNumId = overrides.TryGetValue(listId, out var listOverride) && listOverride.ListId > 0
                ? listOverride.ListId
                : listId;

            if (definitions.ContainsKey(abstractNumId))
            {
                continue;
            }

            definitions[abstractNumId] = CreateFallbackDefinition(abstractNumId);
        }

        var instanceIds = usedListIds.Count > 0
            ? usedListIds
            : overrides.Keys.Union(definitions.Keys).OrderBy(id => id).ToList();

        var instances = new List<NumberingInstanceDefinition>();
        foreach (var numId in instanceIds)
        {
            ListFormatOverride? listOverride = null;
            var abstractNumId = overrides.TryGetValue(numId, out listOverride) && listOverride.ListId > 0
                ? listOverride.ListId
                : numId;

            if (!definitions.ContainsKey(abstractNumId))
            {
                definitions[abstractNumId] = CreateFallbackDefinition(abstractNumId);
            }

            instances.Add(new NumberingInstanceDefinition
            {
                NumId = numId,
                AbstractNumId = abstractNumId,
                LevelOverrides = listOverride != null
                    ? GetEffectiveLevelOverrides(definitions[abstractNumId], listOverride)
                    : new List<ListLevelOverride>()
            });
        }

        return new NumberingPackage
        {
            AbstractDefinitions = definitions.Values.OrderBy(definition => definition.Id).ToList(),
            Instances = instances.OrderBy(instance => instance.NumId).ToList()
        };
    }

    private static List<ListLevelOverride> GetEffectiveLevelOverrides(NumberingDefinition definition, ListFormatOverride listOverride)
    {
        var baseStarts = definition.Levels.ToDictionary(level => level.Level, level => level.Start);
        var effectiveOverrides = new List<ListLevelOverride>();

        foreach (var levelOverride in listOverride.Levels.OrderBy(level => level.Level))
        {
            if (levelOverride.StartAt <= 0)
            {
                continue;
            }

            if (baseStarts.TryGetValue(levelOverride.Level, out var baseStart) && baseStart == levelOverride.StartAt)
            {
                continue;
            }

            effectiveOverrides.Add(new ListLevelOverride
            {
                Level = levelOverride.Level,
                StartAt = levelOverride.StartAt
            });
        }

        return effectiveOverrides;
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
            WriteLevel(new NumberingLevel
            {
                Level = lvl,
                NumberFormat = NumberFormat.Decimal,
                Text = $"%{lvl + 1}.",
                Start = 1
            });
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
                WriteLevel(level);
            }
        }
        else
        {
            // Fallback: write a single bullet level
            WriteLevel(new NumberingLevel
            {
                Level = 0,
                NumberFormat = NumberFormat.Bullet,
                Text = "\u00B7",
                Start = 1
            });
        }

        _writer.WriteEndElement(); // w:abstractNum
    }

    private void WriteLevel(NumberingLevel level)
    {
        var indentLeft = level.ParagraphProperties?.IndentLeft;
        var hanging = level.ParagraphProperties?.IndentFirstLine < 0
            ? Math.Abs(level.ParagraphProperties.IndentFirstLine)
            : 360;

        _writer.WriteStartElement("w", "lvl", WNs);
        _writer.WriteAttributeString("w", "ilvl", WNs, level.Level.ToString());

        _writer.WriteStartElement("w", "start", WNs);
        _writer.WriteAttributeString("w", "val", WNs, level.Start.ToString());
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "numFmt", WNs);
        _writer.WriteAttributeString("w", "val", WNs, GetNumberFormatValue(level.NumberFormat));
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlText", WNs);
        _writer.WriteAttributeString("w", "val", WNs, level.Text ?? $"%{level.Level + 1}.");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "lvlJc", WNs);
        _writer.WriteAttributeString("w", "val", WNs, "left");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "pPr", WNs);
        _writer.WriteStartElement("w", "ind", WNs);
        _writer.WriteAttributeString("w", "left", WNs, (indentLeft.GetValueOrDefault(720 + level.Level * 720)).ToString());
        _writer.WriteAttributeString("w", "hanging", WNs, hanging.ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement(); // w:pPr

        if (level.RunProperties != null && RunPropertiesHelper.HasRunProperties(level.RunProperties))
        {
            RunPropertiesHelper.WriteStyleRunProperties(_writer, level.RunProperties);
        }

        _writer.WriteEndElement(); // w:lvl
    }

    private void WriteNum(NumberingInstanceDefinition instance)
    {
        _writer.WriteStartElement("w", "num", WNs);
        _writer.WriteAttributeString("w", "numId", WNs, instance.NumId.ToString());

        _writer.WriteStartElement("w", "abstractNumId", WNs);
        _writer.WriteAttributeString("w", "val", WNs, instance.AbstractNumId.ToString());
        _writer.WriteEndElement();

        foreach (var levelOverride in instance.LevelOverrides)
        {
            _writer.WriteStartElement("w", "lvlOverride", WNs);
            _writer.WriteAttributeString("w", "ilvl", WNs, levelOverride.Level.ToString());

            _writer.WriteStartElement("w", "startOverride", WNs);
            _writer.WriteAttributeString("w", "val", WNs, levelOverride.StartAt.ToString());
            _writer.WriteEndElement();

            _writer.WriteEndElement();
        }

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
