#nullable enable
using System.IO;
using System.Text;
using System.Xml;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.Writers;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class ThemeReaderTests
{
    [Fact]
    public void ParseThemeMetadata_ExtractsColorsAndFonts()
    {
        var theme = new ThemeModel
        {
            XmlContent = """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                  <a:themeElements>
                    <a:clrScheme name="Custom Theme">
                      <a:dk1><a:srgbClr val="111111" /></a:dk1>
                      <a:lt1><a:sysClr val="window" lastClr="FFFFFF" /></a:lt1>
                      <a:accent1><a:srgbClr val="4472C4" /></a:accent1>
                      <a:hlink><a:scrgbClr r="0" g="0" b="100000" /></a:hlink>
                    </a:clrScheme>
                    <a:fontScheme name="Custom Fonts">
                      <a:majorFont>
                        <a:latin typeface="Aptos Display" />
                        <a:ea typeface="Microsoft YaHei" />
                        <a:cs typeface="Times New Roman" />
                      </a:majorFont>
                      <a:minorFont>
                        <a:latin typeface="Aptos" />
                        <a:ea typeface="SimSun" />
                        <a:cs typeface="Arial" />
                      </a:minorFont>
                    </a:fontScheme>
                  </a:themeElements>
                </a:theme>
                """
        };

        ThemeReader.ParseThemeMetadata(theme);

        Assert.Equal("4472C4", theme.ColorMap["accent1"]);
        Assert.Equal("0000FF", theme.ColorMap["hlink"]);
        Assert.Equal("Aptos", theme.MinorLatinFont);
        Assert.Equal("SimSun", theme.MinorEastAsiaFont);
        Assert.Equal("Times New Roman", theme.MajorBidiFont);
    }

    [Fact]
    public void StylesWriter_UsesResolvedThemeColorForThemeBackedRunColor()
    {
      var document = new DocumentModel
      {
        Theme = new ThemeModel(),
        Styles = new StyleSheet
        {
          Styles =
          {
            new StyleDefinition
            {
              StyleId = 42,
              Name = "ThemeAccent",
              Type = StyleType.Character,
              RunProperties = new RunProperties { Color = 0x01000000 | 4 }
            }
          }
        }
      };
      document.Theme.ColorMap["accent1"] = "4472C4";

        using var ms = new MemoryStream();
        using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
      new StylesWriter(writer).WriteStyles(document);
        writer.Flush();

        var xml = Encoding.UTF8.GetString(ms.ToArray());
        Assert.Contains("themeColor=\"accent1\"", xml);
        Assert.Contains("val=\"4472C4\"", xml);
    }

    [Fact]
    public void StylesWriter_UsesThemeFontsInDocumentDefaults()
    {
        var document = new DocumentModel
        {
            Theme = new ThemeModel
            {
                MinorLatinFont = "Aptos",
                MinorEastAsiaFont = "SimSun",
                MinorBidiFont = "Arial"
            }
        };

        using var ms = new MemoryStream();
        using var writer = XmlWriter.Create(ms, new XmlWriterSettings { Encoding = Encoding.UTF8 });
        new StylesWriter(writer).WriteStyles(document);
        writer.Flush();

        var xml = Encoding.UTF8.GetString(ms.ToArray());
        Assert.Contains("ascii=\"Aptos\"", xml);
        Assert.Contains("eastAsia=\"SimSun\"", xml);
        Assert.Contains("cs=\"Arial\"", xml);
    }
}