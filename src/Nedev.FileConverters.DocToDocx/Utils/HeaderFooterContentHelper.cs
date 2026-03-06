using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Utils;

internal static class HeaderFooterContentHelper
{
    public static bool HasUsableContent(HeaderFooterModel? headerFooter)
    {
        if (headerFooter == null)
            return false;

        return HasUsableParagraphs(headerFooter.Paragraphs) || HasUsableText(headerFooter.Text);
    }

    public static bool HasUsableParagraphs(IReadOnlyList<ParagraphModel>? paragraphs)
    {
        if (paragraphs == null || paragraphs.Count == 0)
            return false;

        var combinedText = string.Concat(paragraphs.SelectMany(p => p.Runs ?? Enumerable.Empty<RunModel>())
            .Select(r => r.Text));
        if (!string.IsNullOrEmpty(combinedText) && !HasUsableText(combinedText))
            return false;

        int visibleChars = 0;
        foreach (var paragraph in paragraphs)
        {
            if (paragraph.Runs == null)
                continue;

            foreach (var run in paragraph.Runs)
            {
                if (run.IsPicture || run.IsField)
                    return true;

                if (!string.IsNullOrWhiteSpace(run.Text))
                    visibleChars += run.Text.Count(ch => !char.IsWhiteSpace(ch));
            }
        }

        return visibleChars > 0 && visibleChars <= 256;
    }

    public static bool HasUsableText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        int visibleChars = text.Count(ch => !char.IsWhiteSpace(ch));
        return visibleChars > 0 && visibleChars <= 256 && !LooksLikeMixedScriptNoise(text);
    }

    private static bool LooksLikeMixedScriptNoise(string text)
    {
        bool hasLatinOrDigit = false;
        bool hasCjk = false;
        bool hasHangul = false;
        bool hasKana = false;
        bool hasOtherScript = false;

        foreach (var ch in text)
        {
            if (char.IsWhiteSpace(ch) || char.IsPunctuation(ch))
                continue;

            if (char.IsAsciiLetterOrDigit(ch))
            {
                hasLatinOrDigit = true;
                continue;
            }

            if (ch is >= '\u4E00' and <= '\u9FFF' or >= '\u3400' and <= '\u4DBF' or >= '\uF900' and <= '\uFAFF')
            {
                hasCjk = true;
                continue;
            }

            if (ch is >= '\uAC00' and <= '\uD7AF' or >= '\u1100' and <= '\u11FF' or >= '\u3130' and <= '\u318F')
            {
                hasHangul = true;
                continue;
            }

            if (ch is >= '\u3040' and <= '\u30FF' or >= '\u31F0' and <= '\u31FF')
            {
                hasKana = true;
                continue;
            }

            hasOtherScript = true;
        }

        int nonLatinScriptCount = 0;
        if (hasCjk) nonLatinScriptCount++;
        if (hasHangul) nonLatinScriptCount++;
        if (hasKana) nonLatinScriptCount++;
        if (hasOtherScript) nonLatinScriptCount++;

        return !hasLatinOrDigit && ((hasCjk && hasHangul) || nonLatinScriptCount >= 3);
    }
}