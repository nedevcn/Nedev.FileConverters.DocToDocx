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

    [Fact]
    public void ApplyBookmarkMarkers_PreservesBookmarkOnlyEmptyParagraph()
    {
        var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
        var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

        using var docReader = new DocReader(inputPath);
        docReader.Document.Bookmarks.Add(new BookmarkModel
        {
            Name = "emptyMark",
            StartCp = 10,
            EndCp = 10
        });

        var method = typeof(DocReader).GetMethod("ApplyBookmarkMarkers", BindingFlags.Instance | BindingFlags.NonPublic);
        Assert.NotNull(method);

        var annotatedRuns = (List<RunModel>)method!.Invoke(docReader, new object[] { new List<RunModel>(), 10, 10 })!;

        Assert.Collection(
            annotatedRuns,
            run =>
            {
                Assert.True(run.IsBookmark);
                Assert.True(run.IsBookmarkStart);
                Assert.Equal("emptyMark", run.BookmarkName);
                Assert.Equal(10, run.CharacterPosition);
            },
            run =>
            {
                Assert.True(run.IsBookmark);
                Assert.False(run.IsBookmarkStart);
                Assert.Equal("emptyMark", run.BookmarkName);
                Assert.Equal(10, run.CharacterPosition);
            });
    }

    [Fact]
    public void MergeTextboxShapesIntoTextboxes_CopiesRecoveredLayout_AndRemovesDuplicateShapes()
    {
        var document = new DocumentModel();
        document.Textboxes.Add(new TextboxModel
        {
            Index = 1,
            AnchorParagraphIndex = 0,
            Paragraphs =
            {
                new ParagraphModel
                {
                    Properties = new ParagraphProperties { Alignment = ParagraphAlignment.Center },
                    Runs = { new RunModel { Text = "textbox" } }
                }
            }
        });

        document.Shapes.Add(new ShapeModel
        {
            Id = 11,
            Type = ShapeType.Textbox,
            Anchor = new ShapeAnchor
            {
                IsFloating = true,
                X = 120,
                Y = 240,
                Width = 3600,
                Height = 1800,
                WrapType = ShapeWrapType.Through
            }
        });

        DocReader.MergeTextboxShapesIntoTextboxes(document);

        var textbox = Assert.Single(document.Textboxes);
        Assert.Equal(120, textbox.Left);
        Assert.Equal(240, textbox.Top);
        Assert.Equal(3600, textbox.Width);
        Assert.Equal(1800, textbox.Height);
        Assert.Equal(TextboxWrapMode.Through, textbox.WrapMode);
        Assert.Equal(TextboxHorizontalAlignment.Center, textbox.HorizontalAlignment);
        Assert.Empty(document.Shapes);
    }

    [Fact]
    public void MergeTextboxShapesIntoTextboxes_PrefersMatchingAnchorParagraphOverInputOrder()
    {
        var document = new DocumentModel();
        document.Textboxes.Add(new TextboxModel { Index = 1, AnchorParagraphIndex = 4, Paragraphs = { new ParagraphModel { Runs = { new RunModel { Text = "late" } } } } });
        document.Textboxes.Add(new TextboxModel { Index = 2, AnchorParagraphIndex = 1, Paragraphs = { new ParagraphModel { Runs = { new RunModel { Text = "early" } } } } });

        document.Shapes.Add(new ShapeModel
        {
            Id = 101,
            Type = ShapeType.Textbox,
            Anchor = new ShapeAnchor { IsFloating = true, ParagraphIndex = 1, X = 10, Y = 20, Width = 1000, Height = 500, WrapType = ShapeWrapType.Square }
        });
        document.Shapes.Add(new ShapeModel
        {
            Id = 102,
            Type = ShapeType.Textbox,
            Anchor = new ShapeAnchor { IsFloating = true, ParagraphIndex = 4, X = 30, Y = 40, Width = 2000, Height = 800, WrapType = ShapeWrapType.Through }
        });

        DocReader.MergeTextboxShapesIntoTextboxes(document);

        Assert.Equal(30, document.Textboxes[0].Left);
        Assert.Equal(TextboxWrapMode.Through, document.Textboxes[0].WrapMode);
        Assert.Equal(10, document.Textboxes[1].Left);
        Assert.Equal(TextboxWrapMode.Square, document.Textboxes[1].WrapMode);
    }

    [Fact]
    public void AttachTextboxAnchorHints_MapsAnchorCpToParagraphIndex()
    {
        var document = new DocumentModel();
        document.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "p0", CharacterPosition = 0, CharacterLength = 2 } } });
        document.Paragraphs.Add(new ParagraphModel { Index = 1, Runs = { new RunModel { Text = "p1", CharacterPosition = 20, CharacterLength = 2 } } });
        document.Textboxes.Add(new TextboxModel { Index = 1, StoryStartCharacterPosition = 100 });
        document.Textboxes.Add(new TextboxModel { Index = 2, StoryStartCharacterPosition = 200 });

        DocReader.AttachTextboxAnchorHints(document, new[]
        {
            new TextboxAnchorFieldInfo { FieldStartCharacterPosition = 5 },
            new TextboxAnchorFieldInfo { FieldStartCharacterPosition = 25 }
        });

        Assert.Equal(5, document.Textboxes[0].AnchorCharacterPosition);
        Assert.Equal(0, document.Textboxes[0].AnchorParagraphIndex);
        Assert.Equal(25, document.Textboxes[1].AnchorCharacterPosition);
        Assert.Equal(1, document.Textboxes[1].AnchorParagraphIndex);
    }

    [Fact]
    public void BuildTextboxAnchorFields_GroupsFieldTripletsFromTextboxPlcPositions()
    {
        const string text = "\u0013 SHAPE \\* MERGEFORMAT \u0014Text box one\u0015 filler \u0013 SHAPE \u0014Text box two\u0015";
        var plcPositions = new[]
        {
            text.IndexOf(FieldReader.FieldStartChar),
            text.IndexOf(FieldReader.FieldSeparatorChar),
            text.IndexOf(FieldReader.FieldEndChar),
            text.LastIndexOf(FieldReader.FieldStartChar),
            text.LastIndexOf(FieldReader.FieldSeparatorChar),
            text.LastIndexOf(FieldReader.FieldEndChar)
        };

        var fields = DocReader.BuildTextboxAnchorFields(text, text.Length, plcPositions);

        Assert.Equal(2, fields.Count);
        Assert.Equal(plcPositions[0], fields[0].FieldStartCharacterPosition);
        Assert.Equal(plcPositions[1], fields[0].FieldSeparatorCharacterPosition);
        Assert.Equal(plcPositions[2], fields[0].FieldEndCharacterPosition);
        Assert.Equal(plcPositions[3], fields[1].FieldStartCharacterPosition);
        Assert.Equal(plcPositions[4], fields[1].FieldSeparatorCharacterPosition);
        Assert.Equal(plcPositions[5], fields[1].FieldEndCharacterPosition);
    }

    private static void AddChpRange(Dictionary<int, ChpBase> map, int start, int end, int fontSize)
    {
        var chp = new ChpBase { FontSize = (byte)fontSize };
        for (int cp = start; cp < end; cp++)
            map[cp] = chp;
    }
}