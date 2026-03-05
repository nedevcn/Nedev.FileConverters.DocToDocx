using System.Xml;
using Nedev.DocToDocx.Models;
using Nedev.DocToDocx.Utils;

namespace Nedev.DocToDocx.Writers;

/// <summary>
/// Shared helper for writing w:rPr (run properties) to OOXML.
/// Used by both DocumentWriter and StylesWriter to avoid code duplication.
/// </summary>
internal static class RunPropertiesHelper
{
    private const string WNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Checks whether any run properties are set that warrant writing a w:rPr element.
    /// </summary>
    public static bool HasRunProperties(RunProperties props)
    {
        return props.IsBold || props.IsBoldCs || props.IsItalic || props.IsItalicCs ||
               props.IsUnderline || props.IsStrikeThrough || props.IsDoubleStrikeThrough ||
               props.IsSmallCaps || props.IsAllCaps || props.IsSuperscript || props.IsSubscript ||
               props.IsHidden || props.IsOutline || props.IsShadow || props.IsEmboss || props.IsImprint ||
               props.FontSize != 24 || props.FontSizeCs != 24 ||
               props.Color != 0 || props.HasRgbColor || props.HighlightColor > 0 ||
               props.Kerning > 0 || props.CharacterSpacingAdjustment != 0 ||
               props.Position != 0 || (props.CharacterScale != 100 && props.CharacterScale > 0) ||
               !props.SnapToGrid ||
               props.Language > 0 || !string.IsNullOrEmpty(props.LanguageAsia) || !string.IsNullOrEmpty(props.LanguageCs) ||
               !string.IsNullOrEmpty(props.FontName);
    }

    /// <summary>
    /// Maps an UnderlineType to its OOXML string value.
    /// </summary>
    public static string GetUnderlineType(UnderlineType type)
    {
        return type switch
        {
            UnderlineType.Single => "single",
            UnderlineType.WordsOnly => "word",
            UnderlineType.Double => "double",
            UnderlineType.Dotted => "dotted",
            UnderlineType.Thick => "thick",
            UnderlineType.Dash => "dash",
            UnderlineType.DotDash => "dotDash",
            UnderlineType.DotDotDash => "dotDotDash",
            UnderlineType.Wave => "wave",
            UnderlineType.ThickWave => "thickWave",
            _ => "none"
        };
    }

    /// <summary>
    /// Writes the full w:rPr element for run-level (document body) usage.
    /// Includes all properties: fonts, formatting, color (RGB + theme), kern, spacing,
    /// size, highlight, underline, vertAlign, position, character scale, snap-to-grid,
    /// and language.
    /// </summary>
    public static void WriteRunProperties(XmlWriter writer, RunProperties props)
    {
        if (!HasRunProperties(props)) return;

        // rPr sequence: rFonts -> b -> bCs -> i -> iCs -> caps -> smallCaps -> strike
        // -> vanish -> outline -> shadow -> emboss -> imprint -> color -> kern
        // -> spacing -> sz -> szCs -> highlight -> u -> vertAlign -> position -> w -> snapToGrid -> lang
        writer.WriteStartElement("w", "rPr", WNs);
        WriteRunPropertiesCore(writer, props, includeExtended: true);
        writer.WriteEndElement(); // w:rPr
    }

    /// <summary>
    /// Writes the w:rPr element for style-level usage.
    /// Includes the core subset of properties typically stored in styles.
    /// </summary>
    public static void WriteStyleRunProperties(XmlWriter writer, RunProperties props)
    {
        if (!HasRunProperties(props)) return;

        writer.WriteStartElement("w", "rPr", WNs);
        WriteRunPropertiesCore(writer, props, includeExtended: false);
        writer.WriteEndElement(); // w:rPr
    }

    /// <summary>
    /// Core implementation shared by both run-level and style-level writing.
    /// When includeExtended is true, additional properties (hidden, outline/shadow/emboss/imprint,
    /// theme color, kern, spacing, position, character scale, snap-to-grid, language) are emitted.
    /// </summary>
    private static void WriteRunPropertiesCore(XmlWriter writer, RunProperties props, bool includeExtended)
    {
        // 1. rFonts
        if (!string.IsNullOrEmpty(props.FontName))
        {
            writer.WriteStartElement("w", "rFonts", WNs);
            writer.WriteAttributeString("w", "ascii", WNs, props.FontName);
            writer.WriteAttributeString("w", "hAnsi", WNs, props.FontName);
            writer.WriteAttributeString("w", "cs", WNs, props.FontName);
            writer.WriteEndElement();
        }

        // 2. b / bCs
        if (props.IsBold)
        {
            writer.WriteStartElement("w", "b", WNs);
            writer.WriteEndElement();
        }
        if (props.IsBoldCs)
        {
            writer.WriteStartElement("w", "bCs", WNs);
            writer.WriteEndElement();
        }

        // 3. i / iCs
        if (props.IsItalic)
        {
            writer.WriteStartElement("w", "i", WNs);
            writer.WriteEndElement();
        }
        if (props.IsItalicCs)
        {
            writer.WriteStartElement("w", "iCs", WNs);
            writer.WriteEndElement();
        }

        // 4. caps / smallCaps
        if (props.IsAllCaps)
        {
            writer.WriteStartElement("w", "caps", WNs);
            writer.WriteEndElement();
        }
        if (props.IsSmallCaps)
        {
            writer.WriteStartElement("w", "smallCaps", WNs);
            writer.WriteEndElement();
        }

        // 5. strike
        if (props.IsStrikeThrough || props.IsDoubleStrikeThrough)
        {
            writer.WriteStartElement("w", "strike", WNs);
            writer.WriteEndElement();
        }

        // 5.5 hidden text
        if (includeExtended && props.IsHidden)
        {
            writer.WriteStartElement("w", "vanish", WNs);
            writer.WriteEndElement();
        }

        // 6. outline / shadow / emboss / imprint
        if (includeExtended)
        {
            if (props.IsOutline)
            {
                writer.WriteStartElement("w", "outline", WNs);
                writer.WriteEndElement();
            }
            if (props.IsShadow)
            {
                writer.WriteStartElement("w", "shadow", WNs);
                writer.WriteEndElement();
            }
            if (props.IsEmboss)
            {
                writer.WriteStartElement("w", "emboss", WNs);
                writer.WriteEndElement();
            }
            if (props.IsImprint)
            {
                writer.WriteStartElement("w", "imprint", WNs);
                writer.WriteEndElement();
            }
        }

        // 7. color
        if (includeExtended && (props.Color != 0 || props.HasRgbColor))
        {
            string? themeColor = ColorHelper.GetThemeColorName(props.Color);
            string colorHex = props.HasRgbColor
                ? ColorHelper.RgbToHex(props.RgbColor)
                : ColorHelper.ColorToHex(props.Color);

            if (themeColor != null || colorHex != "auto")
            {
                writer.WriteStartElement("w", "color", WNs);
                if (themeColor != null)
                {
                    writer.WriteAttributeString("w", "themeColor", WNs, themeColor);
                }
                else
                {
                    writer.WriteAttributeString("w", "val", WNs, colorHex);
                }
                writer.WriteEndElement();
            }
        }
        else if (!includeExtended && props.Color != 0)
        {
            // Style-level: simple ICO index only
            var colorHex = ColorHelper.ColorToHex(props.Color);
            if (colorHex != "auto")
            {
                writer.WriteStartElement("w", "color", WNs);
                writer.WriteAttributeString("w", "val", WNs, colorHex);
                writer.WriteEndElement();
            }
        }

        // 8. kern
        if (includeExtended && props.Kerning > 0)
        {
            writer.WriteStartElement("w", "kern", WNs);
            writer.WriteAttributeString("w", "val", WNs, props.Kerning.ToString());
            writer.WriteEndElement();
        }

        // 9. spacing (character spacing)
        if (includeExtended && props.CharacterSpacingAdjustment != 0)
        {
            writer.WriteStartElement("w", "spacing", WNs);
            writer.WriteAttributeString("w", "val", WNs, props.CharacterSpacingAdjustment.ToString());
            writer.WriteEndElement();
        }

        // 10. sz / szCs
        if (props.FontSize > 0 && props.FontSize != 24)
        {
            writer.WriteStartElement("w", "sz", WNs);
            writer.WriteAttributeString("w", "val", WNs, props.FontSize.ToString());
            writer.WriteEndElement();
        }
        if (props.FontSizeCs > 0 && props.FontSizeCs != 24)
        {
            writer.WriteStartElement("w", "szCs", WNs);
            writer.WriteAttributeString("w", "val", WNs, props.FontSizeCs.ToString());
            writer.WriteEndElement();
        }

        // 11. highlight
        if (props.HighlightColor > 0)
        {
            writer.WriteStartElement("w", "highlight", WNs);
            writer.WriteAttributeString("w", "val", WNs, ColorHelper.GetHighlightName(props.HighlightColor));
            writer.WriteEndElement();
        }

        // 12. u
        if (props.IsUnderline)
        {
            writer.WriteStartElement("w", "u", WNs);
            writer.WriteAttributeString("w", "val", WNs, GetUnderlineType(props.UnderlineType));
            writer.WriteEndElement();
        }

        // 13. vertAlign
        if (props.IsSuperscript)
        {
            writer.WriteStartElement("w", "vertAlign", WNs);
            writer.WriteAttributeString("w", "val", WNs, "superscript");
            writer.WriteEndElement();
        }
        else if (props.IsSubscript)
        {
            writer.WriteStartElement("w", "vertAlign", WNs);
            writer.WriteAttributeString("w", "val", WNs, "subscript");
            writer.WriteEndElement();
        }

        if (includeExtended)
        {
            // Explicit position offset (in half-points)
            if (props.Position != 0 && !props.IsSuperscript && !props.IsSubscript)
            {
                writer.WriteStartElement("w", "position", WNs);
                writer.WriteAttributeString("w", "val", WNs, props.Position.ToString());
                writer.WriteEndElement();
            }

            // Character Scale (w)
            if (props.CharacterScale != 100 && props.CharacterScale > 0)
            {
                writer.WriteStartElement("w", "w", WNs);
                writer.WriteAttributeString("w", "val", WNs, props.CharacterScale.ToString());
                writer.WriteEndElement();
            }

            // Snap to Grid
            if (!props.SnapToGrid)
            {
                writer.WriteStartElement("w", "snapToGrid", WNs);
                writer.WriteAttributeString("w", "val", WNs, "0");
                writer.WriteEndElement();
            }

            // lang
            if (props.Language > 0 || !string.IsNullOrEmpty(props.LanguageAsia) || !string.IsNullOrEmpty(props.LanguageCs))
            {
                writer.WriteStartElement("w", "lang", WNs);
                if (props.Language > 0)
                {
                    var lang = props.Language switch
                    {
                        0x0409 => "en-US",
                        0x0804 => "zh-CN",
                        0x0404 => "zh-TW",
                        0x0411 => "ja-JP",
                        0x0412 => "ko-KR",
                        0x0407 => "de-DE",
                        0x040C => "fr-FR",
                        0x0410 => "it-IT",
                        0x0C0A => "es-ES",
                        _ => null
                    };
                    if (lang != null)
                    {
                        writer.WriteAttributeString("w", "val", WNs, lang);
                    }
                }
                if (!string.IsNullOrEmpty(props.LanguageAsia))
                {
                    writer.WriteAttributeString("w", "eastAsia", WNs, props.LanguageAsia);
                }
                if (!string.IsNullOrEmpty(props.LanguageCs))
                {
                    writer.WriteAttributeString("w", "bidi", WNs, props.LanguageCs);
                }
                writer.WriteEndElement();
            }
        }
    }
}
