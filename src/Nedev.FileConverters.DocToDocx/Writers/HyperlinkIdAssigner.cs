using Nedev.FileConverters.DocToDocx.Models;

namespace Nedev.FileConverters.DocToDocx.Writers;

/// <summary>
/// Assigns deterministic relationship IDs to external hyperlinks in the document model.
/// This must run before writing document.xml and document.xml.rels so both sides agree on IDs.
/// </summary>
public static class HyperlinkIdAssigner
{
    public static void AssignHyperlinkIds(DocumentModel document, int startId)
    {
        var currentId = startId;

        // Build a map from (Url, Bookmark) to relationship ID so duplicates share one rel
        var relIdByKey = new Dictionary<(string UrlKey, string? Bookmark), string>();

        // Helper to register a hyperlink model
        void RegisterHyperlink(HyperlinkModel hyperlink)
        {
            if (!hyperlink.IsExternal || string.IsNullOrEmpty(hyperlink.Url))
                return;

            var key = (hyperlink.Url.ToLowerInvariant(), hyperlink.Bookmark);
            if (!relIdByKey.TryGetValue(key, out var relId))
            {
                relId = $"rId{currentId++}";
                relIdByKey[key] = relId;
            }

            hyperlink.RelationshipId = relId;
        }

        // First, ensure all HyperlinkModel instances have IDs
        foreach (var hyperlink in document.Hyperlinks)
        {
            RegisterHyperlink(hyperlink);
        }

        // Then, walk all runs and assign/propagate IDs based on their URLs
        void ProcessRuns(IEnumerable<RunModel> runs)
        {
            foreach (var run in runs)
            {
                if (!run.IsHyperlink || string.IsNullOrEmpty(run.HyperlinkUrl))
                    continue;

                // split url and optional fragment so we dedupe the same way as HyperlinkModel
                string original = run.HyperlinkUrl!;
                string url = original;
                string? bookmark = null;
                int hash = url.IndexOf('#');
                if (hash >= 0)
                {
                    bookmark = url.Substring(hash + 1);
                    url = url.Substring(0, hash);
                }

                var key = (url.ToLowerInvariant(), bookmark);
                if (!relIdByKey.TryGetValue(key, out var relId))
                {
                    relId = $"rId{currentId++}";
                    relIdByKey[key] = relId;
                }

                run.HyperlinkRelationshipId = relId;

                // ensure hyperlink model exists so relationships writer will emit it
                if (document.Hyperlinks == null)
                    document.Hyperlinks = new List<HyperlinkModel>();

                bool exists = document.Hyperlinks.Any(h =>
                {
                    string hurl = h.Url ?? string.Empty;
                    int hhash = hurl.IndexOf('#');
                    if (hhash >= 0)
                        hurl = hurl.Substring(0, hhash);
                    return string.Equals(hurl, url, StringComparison.OrdinalIgnoreCase) &&
                           h.Bookmark == bookmark;
                });

                if (!exists)
                {
                    document.Hyperlinks.Add(new HyperlinkModel
                    {
                        Url = original,   // preserve fragment in the stored url
                        Bookmark = bookmark,
                        IsExternal = true,
                        RelationshipId = relId
                    });
                }
            }
        }

        // Body paragraphs
        foreach (var paragraph in document.Paragraphs)
        {
            ProcessRuns(paragraph.Runs);
        }

        // Footnotes
        foreach (var note in document.Footnotes)
        {
            foreach (var paragraph in note.Paragraphs)
            {
                ProcessRuns(paragraph.Runs);
            }
        }

        // Endnotes
        foreach (var note in document.Endnotes)
        {
            foreach (var paragraph in note.Paragraphs)
            {
                ProcessRuns(paragraph.Runs);
            }
        }

        // Textboxes
        foreach (var textbox in document.Textboxes)
        {
            ProcessRuns(textbox.Runs);
            foreach (var paragraph in textbox.Paragraphs)
            {
                ProcessRuns(paragraph.Runs);
            }
        }
    }
}

