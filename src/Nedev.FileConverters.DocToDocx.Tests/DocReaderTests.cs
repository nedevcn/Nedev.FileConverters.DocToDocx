#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests;

public class DocReaderTests
{
    [Fact]
    public void ParseRunsInParagraph_PreservesSplitHyperlinkFieldAcrossRuns()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);

        const string paraText = "\u0013HYPERLINK \"http://example.com\"\u0014click\u0015";
        var chpMap = new Dictionary<int, ChpBase>();
        AddChpRange(chpMap, 0, 7, fontSize: 24);
        AddChpRange(chpMap, 7, 31, fontSize: 26);
        AddChpRange(chpMap, 31, 37, fontSize: 28);
        AddChpRange(chpMap, 37, paraText.Length, fontSize: 30);
        var papMap = new Dictionary<int, PapBase>();
        int imageCounter = 0;

        var method = typeof(DocReader).GetMethod("ParseRunsInParagraph", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(method);

        var parameters = new object[] { paraText, 0, chpMap, papMap, imageCounter };
        var runs = (List<RunModel>)method!.Invoke(docReader, parameters)!;
        Assert.Contains(runs, run => string.Equals(run.FieldCode, "HYPERLINK \"http://example.com\"", StringComparison.Ordinal));
        var hyperlinkRun = runs.Single(run => run.IsHyperlink);
        Assert.True(hyperlinkRun.IsHyperlink);
        Assert.Equal("http://example.com", hyperlinkRun.HyperlinkUrl);
        Assert.Contains("click", hyperlinkRun.Text, StringComparison.Ordinal);
        Assert.Contains(docReader.Document.Hyperlinks, hyperlink => string.Equals(hyperlink.Url, "http://example.com", StringComparison.Ordinal));
        Assert.Single(runs.Where(run => run.IsHyperlink));
    }

    [Fact]
    public void ParseRunsInParagraph_DoesNotLeakHyperlinkFieldCodeIntoVisibleResultText()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);

        const string paraText = "\u0013HYPERLINK \"http://example.com\"\u0014click\u0015";
        var chpMap = new Dictionary<int, ChpBase>();
        AddChpRange(chpMap, 0, paraText.Length, fontSize: 24);
        var papMap = new Dictionary<int, PapBase>();
        int imageCounter = 0;

        var method = typeof(DocReader).GetMethod("ParseRunsInParagraph", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(method);

        var parameters = new object[] { paraText, 0, chpMap, papMap, imageCounter };
        var runs = (List<RunModel>)method!.Invoke(docReader, parameters)!;
        Assert.Contains(runs, run => string.Equals(run.FieldCode, "HYPERLINK \"http://example.com\"", StringComparison.Ordinal));
        var hyperlinkRun = runs.Single(run => run.IsHyperlink);
        Assert.Equal("click", hyperlinkRun.Text);
        Assert.Equal("http://example.com", hyperlinkRun.HyperlinkUrl);
        Assert.DoesNotContain(runs.Where(run => !run.IsHyperlink), run => run.Text.Contains("click", StringComparison.Ordinal));
        Assert.DoesNotContain(runs.Where(run => run.IsHyperlink), run => run.Text.Contains("HYPERLINK", StringComparison.Ordinal));
        Assert.DoesNotContain(runs, run => run.Text.Contains("HYPERLINK", StringComparison.Ordinal));
        Assert.Single(runs.Where(run => run.IsHyperlink));
    }

    [Fact]
    public void ApplyBookmarkMarkers_SplitsRunsAndInjectsBookmarkBoundaries()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Document.Bookmarks.Add(new BookmarkModel
        {
            Name = "mark",
            StartCp = 2,
            EndCp = 4
        });

        var method = typeof(DocReader).GetMethod("ApplyBookmarkMarkers", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(method);

        var runs = new List<RunModel>
        {
            new RunModel
            {
                Text = "hello",
                CharacterPosition = 0,
                CharacterLength = 5,
                Properties = new RunProperties { FontSize = 24 }
            }
        };

        var annotatedRuns = (List<RunModel>)method!.Invoke(docReader, new object[] { runs, 0, 5 })!;

        Assert.Collection(
            annotatedRuns,
            run => Assert.Equal("he", run.Text),
            run =>
            {
                Assert.True(run.IsBookmark);
                Assert.True(run.IsBookmarkStart);
                Assert.Equal("mark", run.BookmarkName);
                Assert.Equal(2, run.CharacterPosition);
            },
            run => Assert.Equal("ll", run.Text),
            run =>
            {
                Assert.True(run.IsBookmark);
                Assert.False(run.IsBookmarkStart);
                Assert.Equal("mark", run.BookmarkName);
                Assert.Equal(4, run.CharacterPosition);
            },
            run => Assert.Equal("o", run.Text));
    }

    private static void AddChpRange(Dictionary<int, ChpBase> map, int start, int end, int fontSize)
    {
        var chp = new ChpBase { FontSize = (byte)fontSize };
        for (int cp = start; cp < end; cp++)
            map[cp] = chp;
    }
}