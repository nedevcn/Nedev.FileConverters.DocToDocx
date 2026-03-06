using System.Text;
using System.Text.RegularExpressions;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Field reader - parses Word field codes.
///
/// In Word 97-2003, fields are delimited by special characters:
///   - 0x13: Field start
///   - 0x14: Field separator
///   - 0x15: Field end
///
/// Field structure:
///   { FIELD_CODE [switches] } [result] }
///
/// Common field types:
///   - PAGE: Page number
///   - NUMPAGES: Number of pages
///   - DATE: Current date
///   - TIME: Current time
///   - AUTHOR: Document author
///   - TITLE: Document title
///   - SUBJECT: Document subject
///   - HYPERLINK: Hyperlink (handled separately)
///   - REF: Cross-reference
///   - TOC: Table of contents
///   - INDEX: Index
/// </summary>
public class FieldReader
{
    // Field delimiters
    public const char FieldStartChar = '\x13';
    public const char FieldSeparatorChar = '\x14';
    public const char FieldEndChar = '\x15';

    // Regex to parse field codes
    private static readonly Regex FieldCodeRegex = new(
        @"^(\w+)\s*(.*)$",
        RegexOptions.Compiled);

    /// <summary>
    /// Parses a field code string and returns field information.
    /// </summary>
    public FieldModel? ParseField(string fieldCode)
    {
        if (string.IsNullOrWhiteSpace(fieldCode))
            return null;

        var trimmedCode = fieldCode.Trim();

        // Match field name and parameters
        var match = FieldCodeRegex.Match(trimmedCode);
        if (!match.Success)
            return null;

        var fieldName = match.Groups[1].Value.ToUpperInvariant();
        var fieldArgs = match.Groups[2].Value.Trim();

        var field = new FieldModel
        {
            Type = ParseFieldType(fieldName),
            RawCode = fieldCode,
            FieldName = fieldName,
            Arguments = fieldArgs,
            Switches = ParseSwitches(fieldArgs)
        };

        return field;
    }

    /// <summary>
    /// Determines if text contains field delimiters.
    /// </summary>
    public bool ContainsFields(string text)
    {
        return !string.IsNullOrEmpty(text) &&
               (text.Contains(FieldStartChar) ||
                text.Contains(FieldSeparatorChar) ||
                text.Contains(FieldEndChar));
    }

    /// <summary>
    /// Extracts plain text from field result.
    /// </summary>
    public string ExtractFieldResult(string fieldText)
    {
        if (string.IsNullOrEmpty(fieldText))
            return string.Empty;

        // Find field separator
        var sepIndex = fieldText.IndexOf(FieldSeparatorChar);
        if (sepIndex < 0)
            return fieldText;

        // Find field end
        var endIndex = fieldText.IndexOf(FieldEndChar, sepIndex);
        if (endIndex < 0)
            return fieldText;

        // Extract result between separator and end
        return fieldText.Substring(sepIndex + 1, endIndex - sepIndex - 1);
    }

    /// <summary>
    /// Parses field type from field name.
    /// </summary>
    private FieldType ParseFieldType(string fieldName)
    {
        return fieldName.ToUpperInvariant() switch
        {
            "PAGE" => FieldType.PageNumber,
            "NUMPAGES" => FieldType.NumPages,
            "SECTION" => FieldType.SectionNumber,
            "DATE" => FieldType.Date,
            "TIME" => FieldType.Time,
            "AUTHOR" => FieldType.Author,
            "TITLE" => FieldType.Title,
            "SUBJECT" => FieldType.Subject,
            "KEYWORDS" => FieldType.Keywords,
            "COMMENTS" => FieldType.Comments,
            "FILENAME" => FieldType.FileName,
            "TEMPLATE" => FieldType.Template,
            "HYPERLINK" => FieldType.Hyperlink,
            "REF" => FieldType.Reference,
            "PAGEREF" => FieldType.PageReference,
            "TOC" => FieldType.TableOfContents,
            "INDEX" => FieldType.Index,
            "XE" => FieldType.IndexEntry,
            "TC" => FieldType.TocEntry,
            "SEQ" => FieldType.Sequence,
            "STYLEREF" => FieldType.StyleReference,
            "ASK" => FieldType.Ask,
            "FILLIN" => FieldType.FillIn,
            "MERGEFIELD" => FieldType.MergeField,
            "IF" => FieldType.If,
            "COMPARE" => FieldType.Compare,
            "FORMULA" or "=" or "EQ" => FieldType.Formula,
            "QUOTE" => FieldType.Quote,
            "SYMBOL" => FieldType.Symbol,
            "EMBED" => FieldType.Embed,
            "LINK" => FieldType.Link,
            "INCLUDETEXT" => FieldType.IncludeText,
            "INCLUDEPICTURE" => FieldType.IncludePicture,
            "BOOKMARK" => FieldType.Bookmark,
            "CREATEDATE" => FieldType.CreateDate,
            "SAVEDATE" => FieldType.SaveDate,
            "PRINTDATE" => FieldType.PrintDate,
            "EDITTIME" => FieldType.EditTime,
            "DOCPROPERTY" => FieldType.DocProperty,
            "USERADDRESS" => FieldType.UserAddress,
            "USERINITIALS" => FieldType.UserInitials,
            "USERNAME" => FieldType.UserName,
            _ => FieldType.Unknown
        };
    }

    /// <summary>
    /// Parses field switches from arguments.
    /// </summary>
    private Dictionary<string, string> ParseSwitches(string args)
    {
        var switches = new Dictionary<string, string>();

        if (string.IsNullOrWhiteSpace(args))
            return switches;

        // Match switches: \switch [value]
        var switchRegex = new Regex(@"\\(\w+)\s*(?:""([^""]*)""|'([^']*)'|(\S+))?");
        var matches = switchRegex.Matches(args);

        foreach (Match match in matches)
        {
            var switchName = match.Groups[1].Value;
            var switchValue = match.Groups[2].Success ? match.Groups[2].Value :
                             match.Groups[3].Success ? match.Groups[3].Value :
                             match.Groups[4].Success ? match.Groups[4].Value : string.Empty;

            switches[switchName] = switchValue;
        }

        return switches;
    }

    /// <summary>
    /// Gets the display value for a field based on its type and current context.
    /// </summary>
    public string GetFieldDisplayValue(FieldModel field, DocumentModel document)
    {
        return field.Type switch
        {
            FieldType.PageNumber => "1", // Would need actual page number
            FieldType.NumPages => "1",   // Would need actual page count
            FieldType.Date => DateTime.Now.ToString(GetDateFormat(field.Switches)),
            FieldType.Time => DateTime.Now.ToString(GetTimeFormat(field.Switches)),
            FieldType.Author => document.Properties.Author ?? "",
            FieldType.Title => document.Properties.Title ?? "",
            FieldType.Subject => document.Properties.Subject ?? "",
            FieldType.FileName => document.Properties.FileName ?? "",
            FieldType.CreateDate => document.Properties.Created.ToString(GetDateFormat(field.Switches)),
            FieldType.SaveDate => document.Properties.Modified.ToString(GetDateFormat(field.Switches)),
            FieldType.UserName => Environment.UserName,
            _ => $"[{field.FieldName}]"
        };
    }

    /// <summary>
    /// Gets date format from switches.
    /// </summary>
    private string GetDateFormat(Dictionary<string, string> switches)
    {
        // Check for format switch
        if (switches.TryGetValue("@", out var format) && !string.IsNullOrEmpty(format))
        {
            // Convert Word date format to .NET format
            return ConvertWordDateFormat(format);
        }

        return "yyyy-MM-dd"; // Default format
    }

    /// <summary>
    /// Gets time format from switches.
    /// </summary>
    private string GetTimeFormat(Dictionary<string, string> switches)
    {
        if (switches.TryGetValue("@", out var format) && !string.IsNullOrEmpty(format))
        {
            return ConvertWordDateFormat(format);
        }

        return "HH:mm:ss"; // Default format
    }

    /// <summary>
    /// Converts Word date/time format to .NET format.
    /// </summary>
    private string ConvertWordDateFormat(string wordFormat)
    {
        // Basic conversion - Word uses similar but not identical format codes
        var netFormat = wordFormat
            .Replace("yyyy", "yyyy")
            .Replace("yy", "yy")
            .Replace("MMMM", "MMMM")
            .Replace("MMM", "MMM")
            .Replace("MM", "MM")
            .Replace("M", "M")
            .Replace("dddd", "dddd")
            .Replace("ddd", "ddd")
            .Replace("dd", "dd")
            .Replace("d", "d")
            .Replace("HH", "HH")
            .Replace("H", "H")
            .Replace("hh", "hh")
            .Replace("h", "h")
            .Replace("mm", "mm")
            .Replace("m", "m")
            .Replace("ss", "ss")
            .Replace("s", "s")
            .Replace("tt", "tt");

        return netFormat;
    }
}


