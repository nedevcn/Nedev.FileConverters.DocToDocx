using System.Text;
using System.Text.RegularExpressions;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Utils;

namespace Nedev.FileConverters.DocToDocx.Readers;

/// <summary>
/// Hyperlink reader - extracts hyperlinks from field codes.
///
/// In Word 97-2003, hyperlinks are stored as field codes:
///   HYPERLINK "url" [switches]
///
/// The field structure is:
///   - Field start (19)
///   - Field code (HYPERLINK ...)
///   - Field separator (20)
///   - Display text
///   - Field end (21)
/// </summary>
public class HyperlinkReader
{
    // Regex to match HYPERLINK field codes
    private static readonly Regex HyperlinkRegex = new(
        @"HYPERLINK\s+""([^""]+)""(?:\s+\\l\s+""([^""]+)"")?(?:\s+\\m\s+""([^""]+)"")?",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex HyperlinkSimpleRegex = new(
        @"HYPERLINK\s+(?:""([^""]+)""|'([^']+)'|(\S+))",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public List<HyperlinkModel> Hyperlinks { get; private set; } = new();

    /// <summary>
    /// Parses hyperlinks from a field code string.
    /// </summary>
    public HyperlinkModel? ParseHyperlink(string fieldCode)
    {
        if (string.IsNullOrWhiteSpace(fieldCode))
            return null;

        // Try regex match
        var match = HyperlinkRegex.Match(fieldCode);
        if (!match.Success)
        {
            match = HyperlinkSimpleRegex.Match(fieldCode);
        }

        if (!match.Success)
            return null;

        // Extract URL
        var url = match.Groups[1].Value;
        if (string.IsNullOrEmpty(url) && match.Groups.Count > 2)
        {
            url = match.Groups[2].Value;
        }
        if (string.IsNullOrEmpty(url) && match.Groups.Count > 3)
        {
            url = match.Groups[3].Value;
        }

        if (string.IsNullOrEmpty(url))
            return null;

        // Extract bookmark anchor if present
        var bookmark = match.Groups[2].Success ? match.Groups[2].Value : null;

        return new HyperlinkModel
        {
            Url = url,
            Bookmark = bookmark,
            IsExternal = !url.StartsWith("#") && !string.IsNullOrEmpty(url)
        };
    }

    /// <summary>
    /// Checks if a field code represents a hyperlink.
    /// </summary>
    public bool IsHyperlinkField(string fieldCode)
    {
        return !string.IsNullOrWhiteSpace(fieldCode) &&
               fieldCode.TrimStart().StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Extracts the display text from hyperlink field code if present.
    /// </summary>
    public string GetDisplayText(string fieldCode, string defaultText)
    {
        // If field code has \o switch, it specifies display text
        var match = Regex.Match(fieldCode, @"\\o\s+""([^""]+)""", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            return match.Groups[1].Value;
        }

        return defaultText;
    }

    /// <summary>
    /// Creates a hyperlink model from a URL string.
    /// </summary>
    public HyperlinkModel CreateHyperlink(string url, string? displayText = null)
    {
        return new HyperlinkModel
        {
            Url = url,
            DisplayText = displayText,
            IsExternal = url.StartsWith("http://") ||
                        url.StartsWith("https://") ||
                        url.StartsWith("ftp://") ||
                        url.StartsWith("mailto:")
        };
    }

    /// <summary>
    /// Detects URLs in plain text and converts them to hyperlinks.
    /// </summary>
    public List<HyperlinkModel> DetectUrls(string text)
    {
        var links = new List<HyperlinkModel>();

        // Simple URL detection regex
        var urlRegex = new Regex(
            @"(https?://|ftp://|mailto:)[^\s<>""]+",
            RegexOptions.Compiled);

        var matches = urlRegex.Matches(text);
        foreach (Match match in matches)
        {
            links.Add(new HyperlinkModel
            {
                Url = match.Value,
                IsExternal = true
            });
        }

        return links;
    }
}


