using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// Writes DOCX styles.xml file
/// </summary>
public class StylesWriter
{
    private readonly XmlWriter _writer;
    private DocumentModel? _document;
    
    public StylesWriter(XmlWriter writer)
    {
        _writer = writer;
    }
    
    /// <summary>
    /// Writes the styles.xml content
    /// </summary>
    public void WriteStyles(DocumentModel document)
    {
        _document = document;
        _writer.WriteStartDocument();
        
        // Write root element with namespace
        _writer.WriteStartElement("w", "styles", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        
        // Add XML namespace definitions
        _writer.WriteAttributeString("xmlns", "w", null, "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        _writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        
        // Write document defaults
        WriteDocumentDefaults();
        
        // Write paragraph styles
        WriteParagraphStyles(document);
        
        // Write table styles
        WriteTableStyles(document);
        
        // Write character styles
        WriteCharacterStyles(document);
        
        _writer.WriteEndElement(); // w:styles
        _writer.WriteEndDocument();
    }
    
    private void WriteDocumentDefaults()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "docDefaults", wNs);
        
        // Run defaults
        _writer.WriteStartElement("w", "rPrDefault", wNs);
        _writer.WriteStartElement("w", "rPr", wNs);
        _writer.WriteStartElement("w", "rFonts", wNs);
        _writer.WriteAttributeString("w", "ascii", wNs, "Calibri");
        _writer.WriteAttributeString("w", "eastAsia", wNs, "SimSun");
        _writer.WriteAttributeString("w", "hAnsi", wNs, "Calibri");
        _writer.WriteAttributeString("w", "cs", wNs, "Times New Roman");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "sz", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "24");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "szCs", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "24");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "lang", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "en-US");
        _writer.WriteAttributeString("w", "eastAsia", wNs, "zh-CN");
        _writer.WriteAttributeString("w", "bidi", wNs, "ar-SA");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "pPrDefault", wNs);
        _writer.WriteStartElement("w", "pPr", wNs);
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement(); // w:docDefaults
    }
    
    private void WriteParagraphStyles(DocumentModel document)
    {
        var writtenStyleIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        WriteKnownParagraphStyle(document, 0, "Normal", () => WriteNormalStyle(), writtenStyleIds);

        for (int i = 1; i <= 9; i++)
        {
            var headingId = $"Heading{i}";
            var level = i;
            WriteKnownParagraphStyle(document, level, headingId, () => WriteHeadingStyle(level), writtenStyleIds);
        }

        WriteKnownParagraphStyle(document, null, "Title", () => WriteStyle("Title", "Title", "Normal", 56, true, false), writtenStyleIds);
        WriteKnownParagraphStyle(document, null, "Subtitle", () => WriteStyle("Subtitle", "Subtitle", "Normal", 28, false, true), writtenStyleIds);
        WriteKnownParagraphStyle(document, null, "Quote", () => WriteStyle("Quote", "Quote", "Normal", 22, false, false, true), writtenStyleIds);
        WriteKnownParagraphStyle(document, null, "ListParagraph", () => WriteStyle("ListParagraph", "List Paragraph", "Normal", 22, false, false), writtenStyleIds);
        WriteKnownParagraphStyle(document, null, "NoSpacing", () => WriteStyle("NoSpacing", "No Spacing", "Normal", 22, false, false), writtenStyleIds);

        WriteHeaderFooterStyles();
        writtenStyleIds.Add("Header");
        writtenStyleIds.Add("Footer");

        foreach (var style in document.Styles.Styles.Where(s => s.Type == StyleType.Paragraph))
        {
            var id = StyleHelper.GetParagraphStyleId(style.StyleId, style.Name);
            if (writtenStyleIds.Contains(id))
            {
                continue;
            }

            WriteCustomStyle(style);
            writtenStyleIds.Add(id);
        }

        foreach (var fallbackStyle in EnumerateReferencedParagraphStyleFallbacks(document, writtenStyleIds))
        {
            WriteCustomStyle(fallbackStyle);
            writtenStyleIds.Add(StyleHelper.GetParagraphStyleId(fallbackStyle.StyleId, fallbackStyle.Name));
        }
    }

    private void WriteKnownParagraphStyle(DocumentModel document, int? styleIndex, string styleId, Action writeDefaultStyle, ISet<string> writtenStyleIds)
    {
        if (TryGetParagraphStyleByXmlId(document, styleId, styleIndex, out var style))
        {
            WriteCustomStyle(style!);
        }
        else
        {
            writeDefaultStyle();
        }

        writtenStyleIds.Add(styleId);
    }

    private bool TryGetParagraphStyleByXmlId(DocumentModel document, string styleId, int? preferredStyleIndex, out StyleDefinition? style)
    {
        style = null;

        if (document.Styles?.Styles == null || document.Styles.Styles.Count == 0)
        {
            return false;
        }

        if (preferredStyleIndex.HasValue)
        {
            style = document.Styles.Styles.FirstOrDefault(s =>
                s.Type == StyleType.Paragraph &&
                s.StyleId == preferredStyleIndex.Value &&
                string.Equals(StyleHelper.GetParagraphStyleId(s.StyleId, s.Name), styleId, StringComparison.OrdinalIgnoreCase));
            if (style != null)
            {
                return true;
            }
        }

        style = document.Styles.Styles.FirstOrDefault(s =>
            s.Type == StyleType.Paragraph &&
            string.Equals(StyleHelper.GetParagraphStyleId(s.StyleId, s.Name), styleId, StringComparison.OrdinalIgnoreCase));
        return style != null;
    }

    private IEnumerable<StyleDefinition> EnumerateReferencedParagraphStyleFallbacks(DocumentModel document, ISet<string> existingStyleIds)
    {
        var byStyleIndex = document.Paragraphs
            .Where(p => p.Properties != null && p.Properties.StyleIndex >= 0)
            .GroupBy(p => p.Properties!.StyleIndex)
            .OrderBy(group => group.Key);

        foreach (var group in byStyleIndex)
        {
            var styleId = StyleHelper.GetParagraphStyleId(group.Key, null);
            if (existingStyleIds.Contains(styleId))
            {
                continue;
            }

            var sampleParagraph = group.FirstOrDefault();
            var sampleProperties = sampleParagraph?.Properties;
            var sampleRunProperties = sampleParagraph?.Runs
                .Select(run => run.Properties)
                .FirstOrDefault(props => props != null && RunPropertiesHelper.HasRunProperties(props));

            yield return new StyleDefinition
            {
                StyleId = (ushort)group.Key,
                Name = styleId,
                Type = StyleType.Paragraph,
                ParagraphProperties = sampleProperties == null ? new ParagraphProperties() : CloneParagraphProperties(sampleProperties),
                RunProperties = sampleRunProperties == null ? new RunProperties() : CloneRunProperties(sampleRunProperties)
            };
        }
    }

    private static ParagraphProperties CloneParagraphProperties(ParagraphProperties source)
    {
        return new ParagraphProperties
        {
            StyleIndex = source.StyleIndex,
            Alignment = source.Alignment,
            IndentLeft = source.IndentLeft,
            IndentRight = source.IndentRight,
            IndentFirstLine = source.IndentFirstLine,
            SpaceBefore = source.SpaceBefore,
            SpaceAfter = source.SpaceAfter,
            LineSpacing = source.LineSpacing,
            LineSpacingMultiple = source.LineSpacingMultiple,
            KeepWithNext = source.KeepWithNext,
            KeepTogether = source.KeepTogether,
            PageBreakBefore = source.PageBreakBefore,
            BorderTop = source.BorderTop,
            BorderBottom = source.BorderBottom,
            BorderLeft = source.BorderLeft,
            BorderRight = source.BorderRight,
            Shading = source.Shading,
            ListFormatId = source.ListFormatId,
            ListLevel = source.ListLevel,
            OutlineLevel = source.OutlineLevel,
            NumberFormat = source.NumberFormat,
            NumberText = source.NumberText,
            SnapToGrid = source.SnapToGrid,
            AutoSpaceDe = source.AutoSpaceDe,
            AutoSpaceDn = source.AutoSpaceDn,
            WordWrap = source.WordWrap,
            Kinsoku = source.Kinsoku,
            OverflowPunct = source.OverflowPunct,
            TopLinePunct = source.TopLinePunct
        };
    }

    private static RunProperties CloneRunProperties(RunProperties source)
    {
        return new RunProperties
        {
            FontIndex = source.FontIndex,
            FontName = source.FontName,
            FontSize = source.FontSize,
            FontSizeCs = source.FontSizeCs,
            IsBold = source.IsBold,
            IsBoldCs = source.IsBoldCs,
            IsItalic = source.IsItalic,
            IsItalicCs = source.IsItalicCs,
            IsUnderline = source.IsUnderline,
            UnderlineType = source.UnderlineType,
            IsStrikeThrough = source.IsStrikeThrough,
            IsDoubleStrikeThrough = source.IsDoubleStrikeThrough,
            IsSmallCaps = source.IsSmallCaps,
            IsAllCaps = source.IsAllCaps,
            IsHidden = source.IsHidden,
            IsSuperscript = source.IsSuperscript,
            IsSubscript = source.IsSubscript,
            Color = source.Color,
            BgColor = source.BgColor,
            CharacterSpacingAdjustment = source.CharacterSpacingAdjustment,
            Language = source.Language,
            LanguageAsia = source.LanguageAsia,
            LanguageCs = source.LanguageCs,
            HighlightColor = source.HighlightColor,
            RgbColor = source.RgbColor,
            HasRgbColor = source.HasRgbColor,
            IsOutline = source.IsOutline,
            IsShadow = source.IsShadow,
            IsEmboss = source.IsEmboss,
            IsImprint = source.IsImprint,
            Border = source.Border,
            Kerning = source.Kerning,
            Position = source.Position,
            CharacterScale = source.CharacterScale,
            EastAsianLayoutType = source.EastAsianLayoutType,
            IsEastAsianVertical = source.IsEastAsianVertical,
            IsEastAsianVerticalCompress = source.IsEastAsianVerticalCompress,
            SnapToGrid = source.SnapToGrid,
            RubyText = source.RubyText,
            IsDeleted = source.IsDeleted,
            IsInserted = source.IsInserted,
            AuthorIndexDel = source.AuthorIndexDel,
            AuthorIndexIns = source.AuthorIndexIns,
            DateDel = source.DateDel,
            DateIns = source.DateIns
        };
    }
    
    private void WriteNormalStyle()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "paragraph");
        _writer.WriteAttributeString("w", "default", wNs, "1");
        _writer.WriteAttributeString("w", "styleId", wNs, "Normal");
        
        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "Normal");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "qFormat", wNs);
        _writer.WriteEndElement();
        
        _writer.WriteEndElement(); // w:style
    }
    
    private void WriteHeadingStyle(int level)
    {
        var styleId = $"Heading{level}";
        var name = $"Heading {level}";
        var fontSize = level switch
        {
            1 => 32,
            2 => 26,
            3 => 24,
            4 => 22,
            5 => 22,
            _ => 22
        };
        
        _writer.WriteStartElement("w", "style", null);
        _writer.WriteAttributeString("w", "type", null, "paragraph");
        _writer.WriteAttributeString("w", "styleId", null, styleId);
        
        _writer.WriteStartElement("w", "name", null);
        _writer.WriteAttributeString("w", "val", null, name);
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "basedOn", null);
        _writer.WriteAttributeString("w", "val", null, "Normal");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "qFormat", null);
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "pPr", null);
        _writer.WriteStartElement("w", "spacing", null);
        _writer.WriteAttributeString("w", "before", null, (level == 1 ? 240 : 120).ToString());
        _writer.WriteAttributeString("w", "after", null, "0");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "outlineLvl", null);
        _writer.WriteAttributeString("w", "val", null, (level - 1).ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "rPr", null);
        _writer.WriteStartElement("w", "b", null);
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "bCs", null);
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "sz", null);
        _writer.WriteAttributeString("w", "val", null, fontSize.ToString());
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "szCs", null);
        _writer.WriteAttributeString("w", "val", null, fontSize.ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement();
    }
    
    private void WriteStyle(string styleId, string name, string basedOn, int fontSize, bool bold, bool italic, bool quote = false)
    {
        _writer.WriteStartElement("w", "style", null);
        _writer.WriteAttributeString("w", "type", null, "paragraph");
        _writer.WriteAttributeString("w", "styleId", null, styleId);
        
        _writer.WriteStartElement("w", "name", null);
        _writer.WriteAttributeString("w", "val", null, name);
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "basedOn", null);
        _writer.WriteAttributeString("w", "val", null, basedOn);
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "qFormat", null);
        _writer.WriteEndElement();
        
        if (quote)
        {
            _writer.WriteStartElement("w", "pPr", null);
            _writer.WriteStartElement("w", "ind", null);
            _writer.WriteAttributeString("w", "left", null, "720");
            _writer.WriteAttributeString("w", "right", null, "720");
            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }
        
        _writer.WriteStartElement("w", "rPr", null);
        if (bold)
        {
            _writer.WriteStartElement("w", "b", null);
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "bCs", null);
            _writer.WriteEndElement();
        }
        if (italic)
        {
            _writer.WriteStartElement("w", "i", null);
            _writer.WriteEndElement();
            _writer.WriteStartElement("w", "iCs", null);
            _writer.WriteEndElement();
        }
        _writer.WriteStartElement("w", "sz", null);
        _writer.WriteAttributeString("w", "val", null, fontSize.ToString());
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "szCs", null);
        _writer.WriteAttributeString("w", "val", null, fontSize.ToString());
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement();
    }
    
    private void WriteHeaderFooterStyles()
    {
        // Write Header style
        _writer.WriteStartElement("w", "style", null);
        _writer.WriteAttributeString("w", "type", null, "paragraph");
        _writer.WriteAttributeString("w", "styleId", null, "Header");
        
        _writer.WriteStartElement("w", "name", null);
        _writer.WriteAttributeString("w", "val", null, "Header");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "basedOn", null);
        _writer.WriteAttributeString("w", "val", null, "Normal");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "qFormat", null);
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "rPr", null);
        _writer.WriteStartElement("w", "sz", null);
        _writer.WriteAttributeString("w", "val", null, "20");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "szCs", null);
        _writer.WriteAttributeString("w", "val", null, "20");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement();
        
        // Write Footer style
        _writer.WriteStartElement("w", "style", null);
        _writer.WriteAttributeString("w", "type", null, "paragraph");
        _writer.WriteAttributeString("w", "styleId", null, "Footer");
        
        _writer.WriteStartElement("w", "name", null);
        _writer.WriteAttributeString("w", "val", null, "Footer");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "basedOn", null);
        _writer.WriteAttributeString("w", "val", null, "Normal");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "qFormat", null);
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "rPr", null);
        _writer.WriteStartElement("w", "sz", null);
        _writer.WriteAttributeString("w", "val", null, "20");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "szCs", null);
        _writer.WriteAttributeString("w", "val", null, "20");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement();
    }
    
    private void WriteCustomStyle(StyleDefinition style)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "paragraph");
        _writer.WriteAttributeString("w", "styleId", wNs, StyleHelper.GetParagraphStyleId(style.StyleId, style.Name));
        
        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, style.Name ?? $"Style {style.StyleId}");
        _writer.WriteEndElement();

        if (style.BasedOn.HasValue)
        {
            _writer.WriteStartElement("w", "basedOn", wNs);
            
            // Try to find the name of the base style to generate its ID correctly
            var basedOnStyle = _document?.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Paragraph && s.StyleId == style.BasedOn.Value);
            var basedOnId = StyleHelper.GetParagraphStyleId(style.BasedOn.Value, basedOnStyle?.Name);
            
            _writer.WriteAttributeString("w", "val", wNs, basedOnId);
            _writer.WriteEndElement();
        }
        
        _writer.WriteStartElement("w", "qFormat", wNs);
        _writer.WriteEndElement();

        // Paragraph-level properties for this style (if any)
        if (style.ParagraphProperties != null)
        {
            WriteStyleParagraphProperties(style.ParagraphProperties);
        }

        // Run-level properties for this style (if any)
        if (style.RunProperties != null)
        {
            WriteStyleRunProperties(style.RunProperties);
        }

        _writer.WriteEndElement(); // w:style
    }
    
    private void WriteTableStyles(DocumentModel document)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        
        // Write Table Normal style
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "table");
        _writer.WriteAttributeString("w", "default", wNs, "1");
        _writer.WriteAttributeString("w", "styleId", wNs, "TableNormal");
        
        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "Normal Table");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "tblPr", wNs);
        _writer.WriteStartElement("w", "tblInd", wNs);
        _writer.WriteAttributeString("w", "w", wNs, "0");
        _writer.WriteAttributeString("w", "type", wNs, "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "tblCellMar", wNs);
        _writer.WriteStartElement("w", "top", wNs);
        _writer.WriteAttributeString("w", "w", wNs, "0");
        _writer.WriteAttributeString("w", "type", wNs, "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "left", wNs);
        _writer.WriteAttributeString("w", "w", wNs, "108");
        _writer.WriteAttributeString("w", "type", wNs, "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "bottom", wNs);
        _writer.WriteAttributeString("w", "w", wNs, "0");
        _writer.WriteAttributeString("w", "type", wNs, "dxa");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "right", wNs);
        _writer.WriteAttributeString("w", "w", wNs, "108");
        _writer.WriteAttributeString("w", "type", wNs, "dxa");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        // Write Table Grid style
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "table");
        _writer.WriteAttributeString("w", "styleId", wNs, "TableGrid");
        
        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "Table Grid");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "basedOn", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "TableNormal");
        _writer.WriteEndElement();
        
        _writer.WriteStartElement("w", "tblPr", wNs);
        _writer.WriteStartElement("w", "tblBorders", wNs);
        _writer.WriteStartElement("w", "top", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "single");
        _writer.WriteAttributeString("w", "sz", wNs, "4");
        _writer.WriteAttributeString("w", "color", wNs, "auto");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "left", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "single");
        _writer.WriteAttributeString("w", "sz", wNs, "4");
        _writer.WriteAttributeString("w", "color", wNs, "auto");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "bottom", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "single");
        _writer.WriteAttributeString("w", "sz", wNs, "4");
        _writer.WriteAttributeString("w", "color", wNs, "auto");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "right", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "single");
        _writer.WriteAttributeString("w", "sz", wNs, "4");
        _writer.WriteAttributeString("w", "color", wNs, "auto");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "insideH", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "single");
        _writer.WriteAttributeString("w", "sz", wNs, "4");
        _writer.WriteAttributeString("w", "color", wNs, "auto");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "insideV", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "single");
        _writer.WriteAttributeString("w", "sz", wNs, "4");
        _writer.WriteAttributeString("w", "color", wNs, "auto");
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        _writer.WriteEndElement();
        
        _writer.WriteEndElement();
        
        // Write any custom table styles from document
        var existingTableStyles = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "TableNormal", "TableGrid" };
        foreach (var style in document.Styles.Styles.Where(s => s.Type == StyleType.Table))
        {
            var id = StyleHelper.GetTableStyleId(style.StyleId, style.Name);
            if (!existingTableStyles.Contains(id))
            {
                WriteCustomTableStyle(style);
                existingTableStyles.Add(id);
            }
        }
    }
    
    private void WriteCustomTableStyle(StyleDefinition style)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "table");
        _writer.WriteAttributeString("w", "styleId", wNs, StyleHelper.GetTableStyleId(style.StyleId, style.Name));
        
        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, style.Name ?? $"Table Style {style.StyleId}");
        _writer.WriteEndElement();

        var basedOnId = "TableNormal";
        if (style.BasedOn.HasValue)
        {
            var basedOnStyle = _document?.Styles.Styles.FirstOrDefault(s => s.Type == StyleType.Table && s.StyleId == style.BasedOn.Value);
            basedOnId = StyleHelper.GetTableStyleId(style.BasedOn.Value, basedOnStyle?.Name);
        }

        _writer.WriteStartElement("w", "basedOn", wNs);
        _writer.WriteAttributeString("w", "val", wNs, basedOnId);
        _writer.WriteEndElement();

        if (style.IsPrimary || style.IsQuickStyle)
        {
            _writer.WriteStartElement("w", "qFormat", wNs);
            _writer.WriteEndElement();
        }

        if (style.TableProperties != null)
        {
            WriteStyleTableProperties(style.TableProperties);
        }

        if (style.ParagraphProperties != null)
        {
            WriteStyleParagraphProperties(style.ParagraphProperties);
        }

        if (style.RunProperties != null)
        {
            WriteStyleRunProperties(style.RunProperties);
        }
        
        _writer.WriteEndElement();
    }
    
    private void WriteCharacterStyles(DocumentModel document)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        
        // Write Default Paragraph Font
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "character");
        _writer.WriteAttributeString("w", "default", wNs, "1");
        _writer.WriteAttributeString("w", "styleId", wNs, "DefaultParagraphFont");
        
        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "Default Paragraph Font");
        _writer.WriteEndElement();
        
        _writer.WriteEndElement();
        
        // Write any custom character styles from document
        WriteHyperlinkCharacterStyle();

        var existingCharStyles = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "DefaultParagraphFont", "Hyperlink" };
        foreach (var style in document.Styles.Styles.Where(s => s.Type == StyleType.Character))
        {
            var id = StyleHelper.GetCharacterStyleId(style.StyleId, style.Name);
            if (!existingCharStyles.Contains(id))
            {
                WriteCustomCharacterStyle(style);
                existingCharStyles.Add(id);
            }
        }
    }
    
    private void WriteCustomCharacterStyle(StyleDefinition style)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "character");
        _writer.WriteAttributeString("w", "styleId", wNs, StyleHelper.GetCharacterStyleId(style.StyleId, style.Name));
        
        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, style.Name ?? $"Character Style {style.StyleId}");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "basedOn", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "DefaultParagraphFont");
        _writer.WriteEndElement();

        if (style.RunProperties != null)
        {
            WriteStyleRunProperties(style.RunProperties);
        }

        _writer.WriteEndElement();
    }

    private void WriteHyperlinkCharacterStyle()
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "style", wNs);
        _writer.WriteAttributeString("w", "type", wNs, "character");
        _writer.WriteAttributeString("w", "styleId", wNs, "Hyperlink");

        _writer.WriteStartElement("w", "name", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "Hyperlink");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "basedOn", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "DefaultParagraphFont");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "uiPriority", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "99");
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "unhideWhenUsed", wNs);
        _writer.WriteEndElement();

        _writer.WriteStartElement("w", "rPr", wNs);
        _writer.WriteStartElement("w", "color", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "0563C1");
        _writer.WriteAttributeString("w", "themeColor", wNs, "hyperlink");
        _writer.WriteEndElement();
        _writer.WriteStartElement("w", "u", wNs);
        _writer.WriteAttributeString("w", "val", wNs, "single");
        _writer.WriteEndElement();
        _writer.WriteEndElement();

        _writer.WriteEndElement();
    }

    /// <summary>
    /// Writes w:pPr for a style, using the same mapping as document-level paragraph properties
    /// but without list/numbering.
    /// </summary>
    private void WriteStyleParagraphProperties(ParagraphProperties props)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "pPr", wNs);

        // Borders
        // Note: for styles, we currently omit pBdr/shd here to keep
        // implementation simple; paragraph-level borders/shading are
        // still written at the document level by DocumentWriter.

        // Spacing
        bool hasExplicitLineSpacing = props.LineSpacing != 240 || props.LineSpacingMultiple != 1;
        if (props.SpaceBefore > 0 || props.SpaceAfter > 0 || hasExplicitLineSpacing)
        {
            _writer.WriteStartElement("w", "spacing", wNs);
            if (props.SpaceBefore > 0)
                _writer.WriteAttributeString("w", "before", wNs, props.SpaceBefore.ToString());
            if (props.SpaceAfter > 0)
                _writer.WriteAttributeString("w", "after", wNs, props.SpaceAfter.ToString());
            if (hasExplicitLineSpacing)
            {
                int lineVal = props.LineSpacing;
                string lineRule;
                if (props.LineSpacingMultiple == 1)
                {
                    lineRule = "auto";
                }
                else if (lineVal < 0)
                {
                    lineVal = Math.Abs(lineVal);
                    lineRule = "exact";
                }
                else
                {
                    lineRule = "atLeast";
                }
                _writer.WriteAttributeString("w", "line", wNs, lineVal.ToString());
                _writer.WriteAttributeString("w", "lineRule", wNs, lineRule);
            }
            _writer.WriteEndElement();
        }

        // Indentation
        if (props.IndentLeft != 0 || props.IndentRight != 0 || props.IndentFirstLine != 0)
        {
            _writer.WriteStartElement("w", "ind", wNs);
            if (props.IndentLeft != 0)
                _writer.WriteAttributeString("w", "left", wNs, props.IndentLeft.ToString());
            if (props.IndentRight != 0)
                _writer.WriteAttributeString("w", "right", wNs, props.IndentRight.ToString());

            if (props.IndentFirstLine > 0)
            {
                _writer.WriteAttributeString("w", "firstLine", wNs, props.IndentFirstLine.ToString());
            }
            else if (props.IndentFirstLine < 0)
            {
                _writer.WriteAttributeString("w", "hanging", wNs, Math.Abs(props.IndentFirstLine).ToString());
            }
            _writer.WriteEndElement();
        }

        // Alignment
        if (props.Alignment != ParagraphAlignment.Left)
        {
            _writer.WriteStartElement("w", "jc", wNs);
            var alignment = props.Alignment switch
            {
                ParagraphAlignment.Center => "center",
                ParagraphAlignment.Right => "right",
                ParagraphAlignment.Justify => "both",
                ParagraphAlignment.Distributed => "distribute",
                _ => "left"
            };
            _writer.WriteAttributeString("w", "val", wNs, alignment);
            _writer.WriteEndElement();
        }

        // Outline level
        if (props.OutlineLevel >= 0 && props.OutlineLevel < 9)
        {
            _writer.WriteStartElement("w", "outlineLvl", wNs);
            _writer.WriteAttributeString("w", "val", wNs, props.OutlineLevel.ToString());
            _writer.WriteEndElement();
        }

        _writer.WriteEndElement(); // w:pPr
    }

    private void WriteStyleTableProperties(TableProperties props)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "tblPr", wNs);

        if (props.PreferredWidth > 0)
        {
            _writer.WriteStartElement("w", "tblW", wNs);
            _writer.WriteAttributeString("w", "w", wNs, Math.Clamp(props.PreferredWidth, 1, 31680).ToString());
            _writer.WriteAttributeString("w", "type", wNs, "dxa");
            _writer.WriteEndElement();
        }

        if (props.Alignment != TableAlignment.Left)
        {
            _writer.WriteStartElement("w", "jc", wNs);
            var alignment = props.Alignment switch
            {
                TableAlignment.Center => "center",
                TableAlignment.Right => "right",
                _ => "left"
            };
            _writer.WriteAttributeString("w", "val", wNs, alignment);
            _writer.WriteEndElement();
        }

        if (props.Indent != 0)
        {
            _writer.WriteStartElement("w", "tblInd", wNs);
            _writer.WriteAttributeString("w", "w", wNs, Math.Clamp(props.Indent, -31680, 31680).ToString());
            _writer.WriteAttributeString("w", "type", wNs, "dxa");
            _writer.WriteEndElement();
        }

        if (props.BorderTop != null || props.BorderBottom != null || props.BorderLeft != null || props.BorderRight != null || props.BorderInsideH != null || props.BorderInsideV != null)
        {
            _writer.WriteStartElement("w", "tblBorders", wNs);
            if (props.BorderTop != null) WriteStyleBorder("top", props.BorderTop);
            if (props.BorderBottom != null) WriteStyleBorder("bottom", props.BorderBottom);
            if (props.BorderLeft != null) WriteStyleBorder("left", props.BorderLeft);
            if (props.BorderRight != null) WriteStyleBorder("right", props.BorderRight);
            if (props.BorderInsideH != null) WriteStyleBorder("insideH", props.BorderInsideH);
            if (props.BorderInsideV != null) WriteStyleBorder("insideV", props.BorderInsideV);
            _writer.WriteEndElement();
        }

        if (props.Shading != null)
        {
            WriteStyleShading(props.Shading);
        }

        if (props.CellSpacing > 0)
        {
            int sidePadding = Math.Clamp(props.CellSpacing / 2, 0, 720);
            _writer.WriteStartElement("w", "tblCellMar", wNs);
            WriteStyleMargin("top", 0);
            WriteStyleMargin("left", sidePadding);
            WriteStyleMargin("bottom", 0);
            WriteStyleMargin("right", sidePadding);
            _writer.WriteEndElement();
        }

        _writer.WriteEndElement();
    }

    private void WriteStyleBorder(string position, BorderInfo border)
    {
        if (border.Style == BorderStyle.None) return;

        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", position, wNs);
        _writer.WriteAttributeString("w", "val", wNs, GetBorderStyle(border.Style));
        _writer.WriteAttributeString("w", "sz", wNs, border.Width.ToString());
        _writer.WriteAttributeString("w", "space", wNs, border.Space.ToString());
        _writer.WriteAttributeString("w", "color", wNs, ColorHelper.ColorToHex(border.Color));
        _writer.WriteEndElement();
    }

    private void WriteStyleShading(ShadingInfo shading)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", "shd", wNs);
        _writer.WriteAttributeString("w", "val", wNs, !string.IsNullOrEmpty(shading.PatternVal) ? shading.PatternVal : "clear");
        if (shading.ForegroundColor != 0)
            _writer.WriteAttributeString("w", "color", wNs, ColorHelper.ColorToHex(shading.ForegroundColor));
        _writer.WriteAttributeString("w", "fill", wNs, ColorHelper.ColorToHex(shading.BackgroundColor));
        _writer.WriteEndElement();
    }

    private void WriteStyleMargin(string position, int width)
    {
        const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        _writer.WriteStartElement("w", position, wNs);
        _writer.WriteAttributeString("w", "w", wNs, width.ToString());
        _writer.WriteAttributeString("w", "type", wNs, "dxa");
        _writer.WriteEndElement();
    }

    private static string GetBorderStyle(BorderStyle style)
    {
        return style switch
        {
            BorderStyle.Single => "single",
            BorderStyle.Thick => "thick",
            BorderStyle.Double => "double",
            BorderStyle.Dotted => "dotted",
            BorderStyle.Dashed => "dash",
            BorderStyle.DotDash => "dotDash",
            BorderStyle.DotDotDash => "dotDotDash",
            BorderStyle.Wave => "wave",
            _ => "nil"
        };
    }

    /// <summary>
    /// Writes w:rPr for a style, using the shared RunPropertiesHelper.
    /// </summary>
    private void WriteStyleRunProperties(RunProperties props)
    {
        RunPropertiesHelper.WriteStyleRunProperties(_writer, props);
    }
}
