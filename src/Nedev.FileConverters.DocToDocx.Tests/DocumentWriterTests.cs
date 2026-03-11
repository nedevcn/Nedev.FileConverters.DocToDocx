#nullable enable
using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.IO.Compression;
using System.Threading.Tasks;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.Writers;
using Nedev.FileConverters.DocToDocx.Utils;
using Xunit;

namespace Nedev.FileConverters.DocToDocx.Tests
{
    public class DocumentWriterTests
    {
        private readonly Xunit.Abstractions.ITestOutputHelper _output;

        public DocumentWriterTests(Xunit.Abstractions.ITestOutputHelper output)
        {
            _output = output;
        }
        [Fact]
        public void WriteDocument_MinimalParagraph_EmitsTextRun()
        {
            // Arrange: create a document with one paragraph containing a simple run
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "Hello" });
            doc.Paragraphs.Add(para);

            // Act: write the document to XML in memory
            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = new StreamReader(new MemoryStream(ms.ToArray()), Encoding.UTF8).ReadToEnd();
            }

            // Assert: the run text makes it into the output (xml:space attribute may be present)
            Assert.Contains("Hello", xml);
            Assert.Contains("<w:p", xml); // at least one paragraph element
        }

        [Fact]
        public void WriteDocument_DefaultParagraphSpacing_IsNotForced()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel
            {
                Properties = new ParagraphProperties(),
                Runs = { new RunModel { Text = "Hello" } }
            });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = new StreamReader(new MemoryStream(ms.ToArray()), Encoding.UTF8).ReadToEnd();
            }

            Assert.DoesNotContain("<w:spacing w:line=\"240\" w:lineRule=\"atLeast\"", xml, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteRun_Formats_AreEmitted()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "Fmt" };
            run.Properties = new RunProperties { IsBold = true, IsItalic = true, UnderlineType = UnderlineType.Single };
            para.Runs.Add(run);
            doc.Paragraphs.Add(para);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = new StreamReader(new MemoryStream(ms.ToArray()), Encoding.UTF8).ReadToEnd();
            }

            Assert.Contains("<w:b", xml);
            Assert.Contains("<w:i", xml);
            // underline may or may not be emitted depending on writer logic
            Assert.Contains("Fmt", xml);
        }

        [Fact]
        public void WriteRun_TrackChanges_EmitsInsDel()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run1 = new RunModel { Text = "Added" };
            run1.Properties = new RunProperties { IsInserted = true };
            var run2 = new RunModel { Text = "Removed" };
            run2.Properties = new RunProperties { IsDeleted = true };
            para.Runs.Add(run1);
            para.Runs.Add(run2);
            doc.Paragraphs.Add(para);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<w:ins", xml);
            Assert.Contains("<w:del", xml);
            Assert.Contains("Added", xml);
            Assert.Contains("Removed", xml);
        }

        [Fact]
        public void WriteRun_FieldCodes_AreOutput()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "Default" };
            run.IsField = true;
            run.FieldCode = "ASK Name \"John\""; // ask field with default
            para.Runs.Add(run);
            doc.Paragraphs.Add(para);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("instrText", xml);
            Assert.Contains("ASK Name", xml);
            Assert.Contains("Default", xml);
        }

        [Fact]
        public void WriteDocument_HeaderFooterAndFootnotes_CreateParts()
        {
            var doc = new DocumentModel();

            // header/ footer
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel { Type = HeaderFooterType.HeaderOdd, Text = "HDR" });
            doc.HeadersFooters.Footers.Add(new HeaderFooterModel { Type = HeaderFooterType.FooterOdd, Text = "FTR" });

            // footnote
            var note = new FootnoteModel { Index = 1 };
            note.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "fn" } } });
            doc.Footnotes.Add(note);

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var headerEntry = zip.Entries.FirstOrDefault(entry => entry.FullName.StartsWith("word/header", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
            var footerEntry = zip.Entries.FirstOrDefault(entry => entry.FullName.StartsWith("word/footer", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal));

            Assert.NotNull(headerEntry);
            Assert.NotNull(footerEntry);
            Assert.NotNull(zip.GetEntry("word/footnotes.xml"));

            var hdr = new StreamReader(headerEntry!.Open()).ReadToEnd();
            Assert.Contains("HDR", hdr);
            var ftr = new StreamReader(footerEntry!.Open()).ReadToEnd();
            Assert.Contains("FTR", ftr);
            var fnxml = new StreamReader(zip.GetEntry("word/footnotes.xml").Open()).ReadToEnd();
            Assert.Contains("fn", fnxml);
            // footnote elements should use the w:footnote tag and not fall back to an incorrect
            // default namespace (which previously produced `<w xmlns="footnote" …>`).
            Assert.Matches("<w:footnote[^>]*id=", fnxml);
            Assert.DoesNotContain("xmlns=\"footnote\"", fnxml);
        }

        [Fact]
        public void WriteDocument_Annotations_AreWritten()
        {
            var doc = new DocumentModel();
            var annotation = new AnnotationModel { Id = "1", Author = "Joe" };
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "note" });
            annotation.Paragraphs.Add(para);
            doc.Annotations.Add(annotation);

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var commentsEntry = zip.GetEntry("word/comments.xml");
            Assert.NotNull(commentsEntry);
            var xml = new StreamReader(commentsEntry.Open()).ReadToEnd();
            Assert.Contains("note", xml);
            Assert.Contains("Joe", xml);

            var contentTypes = new StreamReader(zip.GetEntry("[Content_Types].xml").Open()).ReadToEnd();
            Assert.Contains("/word/comments.xml", contentTypes);
            Assert.Contains("application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", contentTypes);
        }

        [Fact]
        public void WriteDocument_NumberingInstances_EmitLevelOverridesOnNum()
        {
            var doc = new DocumentModel();
            doc.NumberingDefinitions.Add(new NumberingDefinition
            {
                Id = 42,
                Levels =
                {
                    new NumberingLevel
                    {
                        Level = 0,
                        NumberFormat = NumberFormat.Decimal,
                        Text = "%1.",
                        Start = 1
                    }
                }
            });
            doc.ListFormatOverrides.Add(new ListFormatOverride
            {
                OverrideId = 7,
                ListId = 42,
                Levels =
                {
                    new ListLevelOverride { Level = 0, StartAt = 4 }
                }
            });
            doc.Paragraphs.Add(new ParagraphModel
            {
                Properties = new ParagraphProperties
                {
                    ListFormatId = 7,
                    ListLevel = 0
                },
                Runs = { new RunModel { Text = "item" } }
            });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            var documentXml = XDocument.Load(zip.GetEntry("word/document.xml")!.Open());
            var numberingXml = XDocument.Load(zip.GetEntry("word/numbering.xml")!.Open());
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            var paragraphNumId = documentXml
                .Descendants(w + "numPr")
                .Elements(w + "numId")
                .Attributes(w + "val")
                .Select(attribute => attribute.Value)
                .FirstOrDefault();

            Assert.Equal("7", paragraphNumId);

            var num = numberingXml
                .Descendants(w + "num")
                .FirstOrDefault(element => string.Equals(element.Attribute(w + "numId")?.Value, "7", StringComparison.Ordinal));

            Assert.NotNull(num);
            Assert.Equal("42", num!.Element(w + "abstractNumId")?.Attribute(w + "val")?.Value);

            var levelOverride = num
                .Elements(w + "lvlOverride")
                .FirstOrDefault(element => string.Equals(element.Attribute(w + "ilvl")?.Value, "0", StringComparison.Ordinal));

            Assert.NotNull(levelOverride);
            Assert.Equal("4", levelOverride!.Element(w + "startOverride")?.Attribute(w + "val")?.Value);
        }

        [Fact]
        public void WriteDocument_Annotations_UseThemeAwareRunProperties()
        {
            var doc = new DocumentModel
            {
                Theme = new ThemeModel()
            };
            doc.Theme.ColorMap["accent1"] = "4472C4";

            var annotation = new AnnotationModel { Id = "1", Author = "Joe" };
            annotation.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "note",
                        Properties = new RunProperties
                        {
                            Color = 0x01000000 | 4,
                            HighlightColor = 7,
                            IsUnderline = true,
                            UnderlineType = UnderlineType.Double,
                            IsSuperscript = true,
                            Border = new BorderInfo { Style = BorderStyle.Single, Width = 4, Space = 0, Color = 0x01000000 | 4 }
                        }
                    }
                }
            });
            doc.Annotations.Add(annotation);

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var xml = new StreamReader(zip.GetEntry("word/comments.xml").Open()).ReadToEnd();

            Assert.Contains("themeColor=\"accent1\"", xml);
            Assert.Contains("val=\"4472C4\"", xml);
            Assert.Contains("<w:highlight", xml);
            Assert.Contains("double", xml);
            Assert.Contains("superscript", xml);
            Assert.Contains("<w:bdr", xml);
        }

        [Fact]
        public void WriteDocument_FootnotesAndEndnotes_UseThemeAwareRunProperties()
        {
            var doc = new DocumentModel
            {
                Theme = new ThemeModel()
            };
            doc.Theme.ColorMap["accent1"] = "4472C4";

            doc.Footnotes.Add(new FootnoteModel
            {
                Index = 1,
                Paragraphs =
                {
                    new ParagraphModel
                    {
                        Runs =
                        {
                            new RunModel
                            {
                                Text = "footnote",
                                Properties = new RunProperties
                                {
                                    Color = 0x01000000 | 4,
                                    HighlightColor = 7,
                                    IsUnderline = true,
                                    UnderlineType = UnderlineType.Wave
                                }
                            }
                        }
                    }
                }
            });

            doc.Endnotes.Add(new EndnoteModel
            {
                Index = 2,
                Paragraphs =
                {
                    new ParagraphModel
                    {
                        Runs =
                        {
                            new RunModel
                            {
                                Text = "endnote",
                                Properties = new RunProperties
                                {
                                    Color = 0x01000000 | 4,
                                    IsSubscript = true,
                                    Border = new BorderInfo { Style = BorderStyle.Single, Width = 4, Space = 0, Color = 0x01000000 | 4 }
                                }
                            }
                        }
                    }
                }
            });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var footnotesXml = new StreamReader(zip.GetEntry("word/footnotes.xml").Open()).ReadToEnd();
            var endnotesXml = new StreamReader(zip.GetEntry("word/endnotes.xml").Open()).ReadToEnd();

            Assert.Contains("themeColor=\"accent1\"", footnotesXml);
            Assert.Contains("wave", footnotesXml);
            Assert.Contains("<w:highlight", footnotesXml);
            Assert.Contains("themeColor=\"accent1\"", endnotesXml);
            Assert.Contains("subscript", endnotesXml);
            Assert.Contains("<w:bdr", endnotesXml);
        }

        [Fact]
        public void WriteDocument_HeaderFooterParagraphs_UseThemeContext_AndDisableExternalHyperlinks()
        {
            var doc = new DocumentModel
            {
                Theme = new ThemeModel()
            };
            doc.Theme.ColorMap["accent1"] = "4472C4";

            doc.HeadersFooters.Headers.Add(new HeaderFooterModel
            {
                Type = HeaderFooterType.HeaderOdd,
                Paragraphs =
                {
                    new ParagraphModel
                    {
                        Runs =
                        {
                            new RunModel
                            {
                                Text = "header link",
                                IsHyperlink = true,
                                HyperlinkUrl = "https://example.com",
                                Properties = new RunProperties
                                {
                                    Color = 0x01000000 | 4,
                                    IsUnderline = true,
                                    UnderlineType = UnderlineType.Single
                                }
                            }
                        }
                    }
                }
            });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var headerXml = new StreamReader(zip.Entries.First(entry => entry.FullName.StartsWith("word/header", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).Open()).ReadToEnd();

            Assert.Contains("themeColor=\"accent1\"", headerXml);
            Assert.Contains("val=\"4472C4\"", headerXml);
            Assert.Contains("header link", headerXml);
            Assert.DoesNotContain("<w:hyperlink", headerXml);
        }

        [Fact]
        public void WriteDocument_HeaderFooterTextFallback_SanitizesXmlUnsafeCharacters()
        {
            var doc = new DocumentModel();
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel
            {
                Type = HeaderFooterType.HeaderOdd,
                Text = " HDR\u0001\uFFFD "
            });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var headerXml = new StreamReader(zip.Entries.First(entry => entry.FullName.StartsWith("word/header", StringComparison.Ordinal) && entry.FullName.EndsWith(".xml", StringComparison.Ordinal)).Open()).ReadToEnd();

            Assert.Contains("xml:space=\"preserve\"", headerXml);
            Assert.Contains("> HDR  </w:t>", headerXml);

            using var reader = XmlReader.Create(new StringReader(headerXml));
            while (reader.Read()) { }
        }

        [Fact]
        public void GeneratedPackage_IsWellFormedXml()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "X" } } });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            foreach (var entry in zip.Entries)
            {
                if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                {
                    using var reader = XmlReader.Create(entry.Open());
                    while (reader.Read()) { /* just advance to detect parse errors */ }
                }
            }
        }

        [Fact]
        public void ValidatePackage_Method_ReturnsExpectedResults()
        {
            // create a valid document and save via converter
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "ok" } } });
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            DocToDocxConverter.SaveDocument(doc, path);
            Assert.True(DocToDocxConverter.ValidatePackage(path, out var msg1), msg1);

            // corrupt the first XML entry by injecting invalid bytes
            using (var archive = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update))
            {
                System.IO.Compression.ZipArchiveEntry entry = null;
                foreach (var e in archive.Entries)
                {
                    if (e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        entry = e;
                        break;
                    }
                }
                Assert.NotNull(entry);
                using var s = entry.Open();
                s.Seek(0, SeekOrigin.Begin);
                // write some invalid xml content
                var bytes = Encoding.UTF8.GetBytes("<w:broken><");
                s.Write(bytes, 0, bytes.Length);
            }
            Assert.False(DocToDocxConverter.ValidatePackage(path, out var msg2));
            Assert.NotNull(msg2);
            File.Delete(path);
        }


        [Fact]
        public void EncryptionHelper_XorRoundTrips()
        {
            byte[] original = { 0, 1, 2, 3, 4, 5, 255 };
            uint key = 0xCAFEBABE;
            var encrypted = EncryptionHelper.DecryptXor(original, key);
            Assert.NotEqual(original, encrypted);
            var decrypted = EncryptionHelper.DecryptXor(encrypted, key);
            Assert.Equal(original, decrypted);
        }

        [Fact]
        public void EncryptionHelper_IsXorEncrypted_DetectsSignatures()
        {
            // if the stream contains a common magic number at the start, return false
            var notEncrypted = new MemoryStream(new byte[] { 0xEC, 0xA5 }); // little-endian 0xA5EC
            Assert.False(EncryptionHelper.IsXorEncrypted(notEncrypted, 0));

            var maybeEncrypted = new MemoryStream(new byte[] { 0x00, 0x00, 0x00 });
            Assert.True(EncryptionHelper.IsXorEncrypted(maybeEncrypted, 0));
        }

        [Fact]
        public void TableModel_NestedProperty_Works()
        {
            var t = new TableModel { Index = 3, ParentTableIndex = 1 };
            Assert.True(t.IsNested);
            Assert.Equal(1, t.ParentTableIndex);

            var top = new TableModel { Index = 5 };
            Assert.False(top.IsNested);
            Assert.Null(top.ParentTableIndex);
        }

        [Fact]
        public void WriteDocument_NestedTable_EmitsNestedTbl()
        {
            var nested = new TableModel { Index = 0 };
            nested.ColumnCount = 1;
            var row = new TableRowModel();
            var cell = new TableCellModel { Index = 0, RowIndex = 0, ColumnIndex = 0 };
            cell.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "inner" } } });
            row.Cells.Add(cell);
            nested.Rows.Add(row);

            var parentTable = new TableModel { Index = 1, ColumnCount = 1 };
            var prow = new TableRowModel();
            var pcell = new TableCellModel { Index = 0, RowIndex = 0, ColumnIndex = 0 };
            var para = new ParagraphModel { Type = ParagraphType.NestedTable, NestedTable = nested };
            pcell.Paragraphs.Add(para);
            prow.Cells.Add(pcell);
            parentTable.Rows.Add(prow);

            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "before" } } });
            // add paragraph containing parentTable via writer path
            doc.Paragraphs.Add(new ParagraphModel { Type = ParagraphType.NestedTable, NestedTable = parentTable });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = new StreamReader(new MemoryStream(ms.ToArray()), Encoding.UTF8).ReadToEnd();
            }

            Assert.Contains("before", xml);
            // should have at least two <w:tbl> entries (parent and nested)
            int count = xml.Split("<w:tbl").Length - 1;
            Assert.True(count >= 2, "Expected at least two tables, got " + count);
            Assert.Contains("inner", xml);

            var xDocument = XDocument.Parse(xml);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var parentCell = xDocument
                .Descendants(w + "tbl")
                .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                .Single()
                .Elements(w + "tr")
                .Single()
                .Elements(w + "tc")
                .Single();

            var childElements = parentCell.Elements().Select(element => element.Name).ToArray();
            Assert.True(childElements.Length >= 2);
            Assert.Equal(w + "tbl", childElements[^2]);
            Assert.Equal(w + "p", childElements[^1]);
        }

        [Fact]
        public void WriteDocument_DeeplyNestedTable_ThreeLevels()
        {
            // build level3
            var level3 = new TableModel { Index = 2, ColumnCount = 1 };
            var r3 = new TableRowModel();
            var c3 = new TableCellModel { Index = 0, RowIndex = 0, ColumnIndex = 0 };
            c3.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "L3" } } });
            r3.Cells.Add(c3);
            level3.Rows.Add(r3);

            // level2 contains level3
            var level2 = new TableModel { Index = 1, ColumnCount = 1 };
            var r2 = new TableRowModel();
            var c2 = new TableCellModel { Index = 0, RowIndex = 0, ColumnIndex = 0 };
            c2.Paragraphs.Add(new ParagraphModel { Type = ParagraphType.NestedTable, NestedTable = level3 });
            r2.Cells.Add(c2);
            level2.Rows.Add(r2);

            // level1 contains level2
            var level1 = new TableModel { Index = 0, ColumnCount = 1 };
            var r1 = new TableRowModel();
            var c1 = new TableCellModel { Index = 0, RowIndex = 0, ColumnIndex = 0 };
            c1.Paragraphs.Add(new ParagraphModel { Type = ParagraphType.NestedTable, NestedTable = level2 });
            r1.Cells.Add(c1);
            level1.Rows.Add(r1);

            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "start" } } });
            doc.Paragraphs.Add(new ParagraphModel { Type = ParagraphType.NestedTable, NestedTable = level1 });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            int count = xml.Split("<w:tbl").Length - 1;
            Assert.True(count >= 3, "expected at least three tables, got " + count);
            Assert.Contains("L3", xml);
        }

        [Fact]
        public void WriteDocument_SmartArtShape_IsEmittedWithText()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "body" } } });
            doc.Shapes.Add(new ShapeModel
            {
                Id = 10,
                Type = ShapeType.SmartArt,
                ParagraphIndexHint = 0,
                Text = "node text"
            });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                new DocumentWriter(writer).WriteDocument(doc);
                writer.Flush();
                xml = new StreamReader(new MemoryStream(ms.ToArray()), Encoding.UTF8).ReadToEnd();
            }

            Assert.Contains("node text", xml);
            // ensure text appears inside a wps:txbx (vector shape text box)
            Assert.Contains("<wps:txbx", xml);
        }

        [Fact]
        public void ZipWriter_FullPackage_HasDocumentEntry()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "A" });
            doc.Paragraphs.Add(para);

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            // Inspect the ZIP for expected entry
            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var entry = zip.GetEntry("word/document.xml");
            Assert.NotNull(entry);

            using var entryStream = entry.Open();
            using var reader = new StreamReader(entryStream, Encoding.UTF8);
            var docXml = reader.ReadToEnd();
            // ensure the text content appears; xml:space attribute may be present
            Assert.Contains("A", docXml);
        }

        [Fact]
        public void CroppingValues_AreClamped()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "img", IsPicture = true, ImageIndex = 0, CropTop = -1000, CropRight = 200000 };
            para.Runs.Add(run);
            doc.Paragraphs.Add(para);
            doc.Images.Add(new ImageModel { WidthEMU = 100, HeightEMU = 100, Data = new byte[] {1} });

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }
            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var xml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();
            // ensure crop attributes not negative or >100000
            Assert.DoesNotMatch("t=-", xml);
            Assert.DoesNotMatch("r=\\\"1[0-9]{6}", xml);
        }

        [Fact]
        public void FirstBodyPictureFlag_ClearedImmediately()
        {
            var doc = new DocumentModel();
            // two images, small size so neither looks full-page
            for (int i = 0; i < 2; i++)
            {
                var para = new ParagraphModel();
                var run = new RunModel { Text = "img", IsPicture = true, ImageIndex = i };
                para.Runs.Add(run);
                doc.Paragraphs.Add(para);
                doc.Images.Add(new ImageModel { WidthEMU = 1, HeightEMU = 1, Data = new byte[] {1} });
            }

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }
            var xml = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read).GetEntry("word/document.xml").Open()).ReadToEnd();
            // ensure first picture not expanded to full page (no large cx values)
            Assert.DoesNotMatch("<wp:extent cx=[0-9]{7,}", xml);
        }

        [Fact]
        public void DuplicateBookmarkNames_GetDistinctIds()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = "A", IsBookmark = true, BookmarkName = "x", IsBookmarkStart = true });
            para.Runs.Add(new RunModel { Text = "B", IsBookmark = true, BookmarkName = "x", IsBookmarkStart = false });
            para.Runs.Add(new RunModel { Text = "C", IsBookmark = true, BookmarkName = "x", IsBookmarkStart = true });
            para.Runs.Add(new RunModel { Text = "D", IsBookmark = true, BookmarkName = "x", IsBookmarkStart = false });
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }
            var xml = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read).GetEntry("word/document.xml").Open()).ReadToEnd();
            var starts = Regex.Matches(xml, "bookmarkStart id=\\\"(\\d+)\\\"");
            var ends = Regex.Matches(xml, "bookmarkEnd id=\\\"(\\d+)\\\"");
            Assert.Equal(starts.Count, ends.Count);
            // all ids should be unique
            var allIds = starts.Cast<Match>().Select(m => m.Groups[1].Value)
                .Concat(ends.Cast<Match>().Select(m => m.Groups[1].Value));
            Assert.Equal(allIds.Distinct().Count(), allIds.Count());
        }

        [Fact]
        public void EmptyBookmarkMarkerRuns_AreSerialized()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel { Text = string.Empty, IsBookmark = true, BookmarkName = "target", IsBookmarkStart = true });
            para.Runs.Add(new RunModel { Text = "Body" });
            para.Runs.Add(new RunModel { Text = string.Empty, IsBookmark = true, BookmarkName = "target", IsBookmarkStart = false });
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            var xml = new StreamReader(new ZipArchive(new MemoryStream(pkg), ZipArchiveMode.Read).GetEntry("word/document.xml")!.Open()).ReadToEnd();
            Assert.Contains("bookmarkStart", xml, StringComparison.Ordinal);
            Assert.Contains("bookmarkEnd", xml, StringComparison.Ordinal);
            Assert.Contains("name=\"target\"", xml, StringComparison.Ordinal);
            Assert.Contains(">Body<", xml, StringComparison.Ordinal);
        }

        [Fact]
        public void SanitizeXmlString_PreservesCjkPunctuation()
        {
            var method = typeof(Nedev.FileConverters.DocToDocx.Writers.DocumentWriter)
                .GetMethod("SanitizeXmlString", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            string input = "渠道代理协议（合同编号：202107010001）";
            string cleaned = (string)method.Invoke(null, new object[] { input })!;
            Assert.Equal(input, cleaned);
        }

        [Fact]
        public void HyperlinkFieldCode_PrefixIsSplitAndSanitized()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            // run contains some normal text, then a stray HYPERLINK field code and part
            // of the url, followed by the visible link text.
            para.Runs.Add(new RunModel
            {
                Text = "foo HYPERLINK \"http://a.com\"bar",
                IsHyperlink = true,
                HyperlinkUrl = "http://a.com"
            });
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            var xml = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read).GetEntry("word/document.xml").Open()).ReadToEnd();

            // the prefix text should appear outside of any hyperlink element
            Assert.Contains("foo ", xml);
            // the hyperlink element should not contain the literal HYPERLINK anymore
            Assert.DoesNotMatch("HYPERLINK", xml);
            // the visible link text 'bar' should still be present inside hyperlink
            Assert.Contains("bar", xml);
        }

        [Fact]
        public void HyperlinkIds_IncludeBookmark()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run1 = new RunModel { Text = "link1", IsHyperlink = true, HyperlinkUrl = "http://a.com#foo" };
            var run2 = new RunModel { Text = "link2", IsHyperlink = true, HyperlinkUrl = "http://a.com#bar" };
            para.Runs.Add(run1);
            para.Runs.Add(run2);
            doc.Paragraphs.Add(para);
            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }
            var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();
            var rels = new StreamReader(zip.GetEntry("word/_rels/document.xml.rels").Open()).ReadToEnd();
            Assert.Contains("docLocation=\"foo\"", documentXml);
            Assert.Contains("docLocation=\"bar\"", documentXml);
            Assert.Contains("Target=\"http://a.com\"", rels);
            // ensure two hyperlink relationships only
            var hyperCount = Regex.Matches(rels, "Type=\\\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\\\"").Count;
            Assert.Equal(2, hyperCount);
        }

        [Fact]
        public void InternalBookmarkHyperlink_WritesAnchorWithoutRelationship()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel
            {
                Text = "jump",
                IsHyperlink = true,
                HyperlinkBookmark = "targetBookmark"
            });
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();
            var rels = new StreamReader(zip.GetEntry("word/_rels/document.xml.rels").Open()).ReadToEnd();

            Assert.Contains("anchor=\"targetBookmark\"", documentXml);
            Assert.DoesNotContain("hyperlink", rels, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HyperlinkRun_UsesThemeAwareRunProperties()
        {
            var doc = new DocumentModel
            {
                Theme = new ThemeModel()
            };
            doc.Theme.ColorMap["accent1"] = "4472C4";

            var para = new ParagraphModel();
            para.Runs.Add(new RunModel
            {
                Text = "jump",
                IsHyperlink = true,
                HyperlinkBookmark = "targetBookmark",
                Properties = new RunProperties
                {
                    Color = 0x01000000 | 4,
                    IsUnderline = true,
                    UnderlineType = UnderlineType.Single
                }
            });
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();

            Assert.Contains("anchor=\"targetBookmark\"", documentXml);
            Assert.Contains("themeColor=\"accent1\"", documentXml);
            Assert.Contains("val=\"4472C4\"", documentXml);
        }

        [Fact]
        public void HyperlinkRun_PreservesTabSeparators()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel
            {
                Text = "Entry\t12",
                IsHyperlink = true,
                HyperlinkBookmark = "tocTarget"
            });
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();

            Assert.Contains("anchor=\"tocTarget\"", documentXml);
            Assert.Contains("<w:tab", documentXml, StringComparison.Ordinal);
            Assert.DoesNotContain(">Entry 12<", documentXml, StringComparison.Ordinal);
        }

        [Fact]
        public void TocStyle_WritesRightAlignedDotLeaderTabStop()
        {
            var doc = new DocumentModel();
            doc.Styles.Styles.Add(new StyleDefinition
            {
                StyleId = 42,
                Name = "toc 1",
                Type = StyleType.Paragraph,
                ParagraphProperties = new ParagraphProperties
                {
                    SpaceBefore = 240,
                    SpaceAfter = 120,
                    LineSpacing = 276
                },
                RunProperties = new RunProperties
                {
                    IsBold = true,
                    FontSize = 20,
                    FontSizeCs = 20
                }
            });

            doc.Paragraphs.Add(new ParagraphModel
            {
                Properties = new ParagraphProperties { StyleIndex = 42 },
                Runs =
                {
                    new RunModel { Text = "Entry\t1", IsHyperlink = true, HyperlinkBookmark = "tocTarget" }
                }
            });

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var stylesXml = new StreamReader(zip.GetEntry("word/styles.xml").Open()).ReadToEnd();

            Assert.Contains("styleId=\"toc1\"", stylesXml, StringComparison.Ordinal);
            Assert.Contains("<w:tabs><w:tab w:val=\"right\" w:leader=\"dot\"", stylesXml, StringComparison.Ordinal);
            Assert.Contains("w:pos=\"9360\"", stylesXml, StringComparison.Ordinal);
        }

        [Fact]
        public void TocBookmarkHyperlink_DoesNotForceHyperlinkCharacterStyle()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            para.Runs.Add(new RunModel
            {
                Text = "Entry\t1",
                IsHyperlink = true,
                HyperlinkBookmark = "_Toc123"
            });
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();

            Assert.Contains("anchor=\"_Toc123\"", documentXml, StringComparison.Ordinal);
            Assert.DoesNotContain("<w:rStyle w:val=\"Hyperlink\" />", documentXml, StringComparison.Ordinal);
            Assert.Contains("<w:noProof", documentXml, StringComparison.Ordinal);
            Assert.Contains("<w:tab", documentXml, StringComparison.Ordinal);
        }

        [Fact]
        public void ParagraphKeepWithNext_IsSerialized()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel
            {
                Properties = new ParagraphProperties { KeepWithNext = true },
                Runs = { new RunModel { Text = "Heading" } }
            });

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();

            Assert.Contains("<w:keepNext", documentXml);
        }

        [Fact]
        public void RunFonts_WriteEastAsiaSlot()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel
            {
                Runs =
                {
                    new RunModel
                    {
                        Text = "中文",
                        Properties = new RunProperties { FontName = "SimSun" }
                    }
                }
            });

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();

            Assert.Contains("w:eastAsia=\"SimSun\"", documentXml);
        }

        [Fact]
        public void CustomTableStyle_WritesTableParagraphAndRunDefaults()
        {
            var doc = new DocumentModel();
            var tableStyleId = 42;
            var tableStyleName = "My Table Style";
            var tableStyleXmlId = Nedev.FileConverters.DocToDocx.Utils.StyleHelper.GetTableStyleId(tableStyleId, tableStyleName);

            doc.Styles.Styles.Add(new StyleDefinition
            {
                StyleId = (ushort)tableStyleId,
                Name = tableStyleName,
                Type = StyleType.Table,
                TableProperties = new TableProperties
                {
                    Alignment = TableAlignment.Center,
                    PreferredWidth = 5000,
                    BorderTop = new BorderInfo { Style = BorderStyle.Single, Width = 4, Color = 1 },
                    BorderInsideH = new BorderInfo { Style = BorderStyle.Single, Width = 4, Color = 1 },
                    Shading = new ShadingInfo { BackgroundColor = 0x00FFFF }
                },
                ParagraphProperties = new ParagraphProperties { Alignment = ParagraphAlignment.Center },
                RunProperties = new RunProperties { FontName = "SimSun", FontSize = 28, IsBold = true }
            });

            var table = new TableModel
            {
                StartParagraphIndex = 0,
                EndParagraphIndex = 0,
                Properties = new TableProperties { StyleIndex = tableStyleId },
                Rows =
                {
                    new TableRowModel
                    {
                        Cells =
                        {
                            new TableCellModel
                            {
                                ColumnIndex = 0,
                                Paragraphs = { new ParagraphModel { Runs = { new RunModel { Text = "Cell" } } } },
                                Properties = new TableCellProperties { Width = 2400 }
                            }
                        }
                    }
                }
            };
            table.RowCount = 1;
            table.ColumnCount = 1;
            doc.Tables.Add(table);
            doc.Paragraphs.Add(new ParagraphModel { Index = 0, Type = ParagraphType.TableCell, Runs = { new RunModel { Text = "Cell" } } });

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read);
            var stylesXml = new StreamReader(zip.GetEntry("word/styles.xml").Open()).ReadToEnd();
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();

            Assert.Contains($"styleId=\"{tableStyleXmlId}\"", stylesXml);
            Assert.Contains("<w:tblPr", stylesXml);
            Assert.Contains("<w:tblBorders", stylesXml);
            Assert.Contains("<w:pPr", stylesXml);
            Assert.Contains("<w:rPr", stylesXml);
            Assert.Contains($"<w:tblStyle w:val=\"{tableStyleXmlId}\"", documentXml);
        }

        [Fact]
        public void HyperlinksCanBeDisabled()
        {
            var doc = new DocumentModel();
            var para = new ParagraphModel();
            var run = new RunModel { Text = "click", IsHyperlink = true, HyperlinkUrl = "http://example.com" };
            para.Runs.Add(run);
            doc.Paragraphs.Add(para);

            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                var options = new Writers.DocumentWriterOptions { EnableHyperlinks = false };
                zw.WriteDocument(doc, options);
                zw.Dispose();
                pkg = ms.ToArray();
            }
            var xml = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read).GetEntry("word/document.xml").Open()).ReadToEnd();
            Assert.DoesNotContain("hyperlink", xml, StringComparison.OrdinalIgnoreCase);
            var rels = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read).GetEntry("word/_rels/document.xml.rels").Open()).ReadToEnd();
            Assert.DoesNotContain("hyperlink", rels, StringComparison.OrdinalIgnoreCase);
            // text still present
            Assert.Contains("click", xml);
        }

        [Fact]
        public void WriteDocument_UsesThemeColorsForBordersShadingAndShapes()
        {
            var doc = new DocumentModel
            {
                Theme = new ThemeModel()
            };
            doc.Theme.ColorMap["accent1"] = "4472C4";
            doc.Theme.ColorMap["accent2"] = "ED7D31";

            doc.Paragraphs.Add(new ParagraphModel
            {
                Index = 0,
                Properties = new ParagraphProperties
                {
                    BorderTop = new BorderInfo { Style = BorderStyle.Single, Width = 4, Space = 0, Color = 0x01000000 | 4 },
                    Shading = new ShadingInfo
                    {
                        ForegroundColor = 0x01000000 | 4,
                        BackgroundColor = 0x01000000 | 5
                    }
                },
                Runs = { new RunModel { Text = "Themed" } }
            });

            doc.Shapes.Add(new ShapeModel
            {
                Id = 1,
                Type = ShapeType.Rectangle,
                ParagraphIndexHint = 0,
                FillColor = 0x01000000 | 4,
                LineColor = 0x01000000 | 5,
                LineWidth = 12700,
                Text = "Box"
            });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("w:themeColor=\"accent1\"", xml);
            Assert.Contains("w:themeFill=\"accent2\"", xml);
            Assert.Contains("w:color=\"4472C4\"", xml);
            Assert.Contains("w:fill=\"ED7D31\"", xml);
            Assert.Contains("<a:schemeClr val=\"accent1\"", xml);
            Assert.Contains("<a:schemeClr val=\"accent2\"", xml);
        }

        [Fact]
        public void WriteDocument_ShapeTextboxText_IsSanitized_AndPreservesMeaningfulWhitespace()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel
            {
                Index = 0,
                Runs = { new RunModel { Text = "body" } }
            });

            doc.Shapes.Add(new ShapeModel
            {
                Id = 1,
                Type = ShapeType.Textbox,
                ParagraphIndexHint = 0,
                Text = " Box\u0001\uFFFD text "
            });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("xml:space=\"preserve\"", xml);
            Assert.Contains("> Box  text </w:t>", xml);
        }

        [Fact]
        public void WriteDocument_CustomGeometry_MultiplePaths_RendersAllPaths()
        {
            var doc = new DocumentModel();
            // add placeholder paragraph so shapes with hint=0 will be placed
            doc.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "body" } } });
            // create geometry with two disjoint rectangles (outer + inner hole)
            var geom = new CustomGeometry
            {
                ViewLeft = 0,
                ViewTop = 0,
                ViewRight = 1000,
                ViewBottom = 1000
            };
            // vertices for outer square
            geom.Vertices.Add(new System.Drawing.Point(0, 0));
            geom.Vertices.Add(new System.Drawing.Point(1000, 0));
            geom.Vertices.Add(new System.Drawing.Point(1000, 1000));
            geom.Vertices.Add(new System.Drawing.Point(0, 1000));
            // vertices for inner square
            geom.Vertices.Add(new System.Drawing.Point(250, 250));
            geom.Vertices.Add(new System.Drawing.Point(750, 250));
            geom.Vertices.Add(new System.Drawing.Point(750, 750));
            geom.Vertices.Add(new System.Drawing.Point(250, 750));

            // outer path
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.MoveTo, VertexIndex = 0 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.LineTo, VertexIndex = 1 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.LineTo, VertexIndex = 2 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.LineTo, VertexIndex = 3 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.Close });
            // separator
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.End });
            // inner path
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.MoveTo, VertexIndex = 4 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.LineTo, VertexIndex = 5 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.LineTo, VertexIndex = 6 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.LineTo, VertexIndex = 7 });
            geom.Segments.Add(new ShapePathSegment { Type = SegmentType.Close });

            doc.Shapes.Add(new ShapeModel
            {
                Id = 1,
                Type = ShapeType.Custom,
                ParagraphIndexHint = 0, // ensure writer emits the shape
                CustomGeometry = geom
            });

            // first, exercise the standalone DocumentWriter (faster) and verify
            // the raw XML contains two paths
            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }
            // count only real <a:path> elements, not the surrounding <a:pathLst> tag
            var pathCount = System.Text.RegularExpressions.Regex.Matches(xml, "<a:path(?!Lst)").Count;
            Assert.Equal(2, pathCount);
            Assert.Contains("<a:close", xml); // ensure close tags were emitted

            // now make sure the same geometry survives the full package writer
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(ms.ToArray()), System.IO.Compression.ZipArchiveMode.Read);
                var docXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();
                var pathCount2 = System.Text.RegularExpressions.Regex.Matches(docXml, "<a:path(?!Lst)").Count;
                Assert.Equal(2, pathCount2);
            }
        }


        [Fact]
        public void WriteDocument_GroupShape_EmitsGrpSpWrapper()
        {
            var doc = new DocumentModel();
            // add a paragraph so shapes with hint=0 will be written
            doc.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "body" } } });
            var child = new ShapeModel
            {
                Id = 2,
                Type = ShapeType.Rectangle,
                FillColor = 0xFF0000,
                ParagraphIndexHint = 0
            };
            var group = new ShapeModel
            {
                Id = 1,
                Type = ShapeType.Group,
                ParagraphIndexHint = 0,
                Children = new List<ShapeModel> { child }
            };
            doc.Shapes.Add(group);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<a:grpSp", xml);
            Assert.Contains("Shape 2", xml); // child docPr name should appear
        }

        [Fact]
        public void WriteDocument_GradientFill_IsWrittenAsGradFill()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "body" } } });
            doc.Paragraphs.Add(new ParagraphModel { Index = 1, Runs = { new RunModel { Text = "body" } } });
            doc.Shapes.Add(new ShapeModel
            {
                Id = 5,
                Type = ShapeType.Rectangle,
                ParagraphIndexHint = 0,
                FillType = FillType.LinearGradient,
                GradientAngle = 5400000, // 90 degrees in DrawingML units
                GradientStops = new List<GradientStop>
                {
                    new GradientStop { Color = 0xFF0000, Position = 0 },
                    new GradientStop { Color = 0x00FF00, Position = 1 }
                }
            });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                var dw = new DocumentWriter(writer);
                dw.WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<a:gradFill", xml);
            Assert.Contains("<a:gs", xml);
            Assert.Contains("pos=\"0\"", xml);
            Assert.Contains("pos=\"100000\"", xml);
        }

        [Fact]
        public void WriteDocument_FirstAndEvenHeaders_EnableSectionSemantics()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "body" } } });
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel { Type = HeaderFooterType.HeaderFirst, Text = "first header" });
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel { Type = HeaderFooterType.HeaderOdd, Text = "odd header" });
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel { Type = HeaderFooterType.HeaderEven, Text = "even header" });
            doc.HeadersFooters.Footers.Add(new HeaderFooterModel { Type = HeaderFooterType.FooterFirst, Text = "first footer" });
            doc.HeadersFooters.Footers.Add(new HeaderFooterModel { Type = HeaderFooterType.FooterOdd, Text = "odd footer" });
            doc.HeadersFooters.Footers.Add(new HeaderFooterModel { Type = HeaderFooterType.FooterEven, Text = "even footer" });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml").Open()).ReadToEnd();
            var settingsXml = new StreamReader(zip.GetEntry("word/settings.xml").Open()).ReadToEnd();

            Assert.Contains("<w:titlePg", documentXml);
            Assert.Contains("w:headerReference", documentXml);
            Assert.Contains("w:type=\"default\"", documentXml);
            Assert.Contains("w:type=\"first\"", documentXml);
            Assert.Contains("w:type=\"even\"", documentXml);
            Assert.Contains("w:footerReference", documentXml);
            Assert.Contains("<w:evenAndOddHeaders", settingsXml);
            Assert.DoesNotContain("<w:updateFields", settingsXml);
        }

        [Fact]
        public void WriteDocument_SectionEndingAtTable_UsesParagraphScopedSectionBreak()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel
            {
                Index = 0,
                Type = ParagraphType.TableCell,
                Runs = { new RunModel { Text = "cell" } }
            });
            doc.Paragraphs.Add(new ParagraphModel
            {
                Index = 1,
                Runs = { new RunModel { Text = "after" } }
            });

            var table = new TableModel
            {
                StartParagraphIndex = 0,
                EndParagraphIndex = 0,
                ColumnCount = 1,
                RowCount = 1,
                Rows =
                {
                    new TableRowModel
                    {
                        Cells =
                        {
                            new TableCellModel
                            {
                                ColumnIndex = 0,
                                Paragraphs = { new ParagraphModel { Runs = { new RunModel { Text = "cell" } } } }
                            }
                        }
                    }
                }
            };
            doc.Tables.Add(table);
            doc.Properties.Sections.Add(new SectionInfo { SectionIndex = 0, StartParagraphIndex = 0, BreakCode = 2 });
            doc.Properties.Sections.Add(new SectionInfo { SectionIndex = 1, StartParagraphIndex = 1, BreakCode = 2 });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml")!.Open()).ReadToEnd();
            var xDocument = XDocument.Parse(documentXml);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var bodyElements = xDocument.Root!.Element(w + "body")!.Elements().ToList();

            Assert.Equal(w + "tbl", bodyElements[0].Name);
            Assert.Equal(w + "p", bodyElements[1].Name);
            Assert.NotNull(bodyElements[1].Element(w + "pPr")?.Element(w + "sectPr"));
            Assert.Equal("after", string.Concat(bodyElements[2].Descendants(w + "t").Select(text => text.Value)));
            Assert.Equal(w + "sectPr", bodyElements[^1].Name);
            Assert.Single(bodyElements.Where(element => element.Name == w + "sectPr"));
        }

        [Fact]
        public void WriteDocument_PerSectionHeaders_GetDistinctPartsAndRelationships()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "first" } } });
            doc.Paragraphs.Add(new ParagraphModel { Index = 1, Runs = { new RunModel { Text = "second" } } });
            doc.Properties.Sections.Add(new SectionInfo { SectionIndex = 0, StartParagraphIndex = 0, BreakCode = 2 });
            doc.Properties.Sections.Add(new SectionInfo { SectionIndex = 1, StartParagraphIndex = 1, BreakCode = 2 });
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel { SectionIndex = 0, Type = HeaderFooterType.HeaderOdd, Text = "section 1 header" });
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel { SectionIndex = 1, Type = HeaderFooterType.HeaderOdd, Text = "section 2 header" });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            var documentXml = new StreamReader(zip.GetEntry("word/document.xml")!.Open()).ReadToEnd();
            var relsXml = new StreamReader(zip.GetEntry("word/_rels/document.xml.rels")!.Open()).ReadToEnd();

            var documentXDoc = XDocument.Parse(documentXml);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var headerRelationshipIds = documentXDoc
                .Descendants(w + "sectPr")
                .Elements(w + "headerReference")
                .Where(reference => (string?)reference.Attribute(w + "type") == "default")
                .Select(reference => (string?)reference.Attribute(r + "id"))
                .Where(id => !string.IsNullOrEmpty(id))
                .ToList();

            Assert.Equal(2, headerRelationshipIds.Count);
            Assert.Equal(2, headerRelationshipIds.Distinct(StringComparer.Ordinal).Count());
            Assert.NotNull(zip.GetEntry("word/header1.xml"));
            Assert.NotNull(zip.GetEntry("word/header2.xml"));
            Assert.Contains("Target=\"header1.xml\"", relsXml, StringComparison.Ordinal);
            Assert.Contains("Target=\"header2.xml\"", relsXml, StringComparison.Ordinal);
        }

        [Fact]
        public void WriteDocument_HeaderPart_WritesLocalImageAndOleRelationships()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "body" } } });
            doc.Images.Add(new ImageModel { WidthEMU = 100, HeightEMU = 100, Data = new byte[] { 1 } });
            doc.OleObjects.Add(new OleObjectModel
            {
                ObjectId = "ole-1",
                ProgId = "Excel.Sheet.8",
                ObjectData = new byte[] { 1, 2, 3 },
                ImageIndex = 0
            });
            doc.HeadersFooters.Headers.Add(new HeaderFooterModel
            {
                Type = HeaderFooterType.HeaderOdd,
                Paragraphs =
                {
                    new ParagraphModel
                    {
                        Runs =
                        {
                            new RunModel { IsPicture = true, ImageIndex = 0 },
                            new RunModel { IsPicture = true, IsOle = true, ImageIndex = 0, OleObjectId = "ole-1", OleProgId = "Excel.Sheet.8" }
                        }
                    }
                }
            });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            var headerXml = new StreamReader(zip.GetEntry("word/header1.xml")!.Open()).ReadToEnd();
            var headerRelsXml = new StreamReader(zip.GetEntry("word/_rels/header1.xml.rels")!.Open()).ReadToEnd();

            Assert.Contains("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", headerXml, StringComparison.Ordinal);
            Assert.Contains("r:embed=\"rId1\"", headerXml, StringComparison.Ordinal);
            Assert.Contains("o:OLEObject", headerXml, StringComparison.Ordinal);
            Assert.Contains("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"", headerRelsXml, StringComparison.Ordinal);
            Assert.Contains("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject\"", headerRelsXml, StringComparison.Ordinal);
        }

        [Fact]
        public void SettingsWriter_WritesCompatibilityFlags()
        {
            var doc = new DocumentModel();
            doc.Properties.FUsePrinterMetrics = true;
            doc.Properties.FSuppressBottomSpacing = true;
            doc.Properties.FSuppressSpacings = true;

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                new SettingsWriter(writer).WriteSettings(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<w:compat", xml);
            Assert.Contains("<w:usePrinterMetrics", xml);
            Assert.Contains("<w:suppressBottomSpacing", xml);
            Assert.Contains("<w:suppressSpacingAtTopOfPage", xml);
        }

        [Fact]
        public void ContentTypes_IncludeExtendedImageFormats()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "img", IsPicture = true, ImageIndex = 0 } } });
            doc.Images.Add(new ImageModel { WidthEMU = 100, HeightEMU = 100, Data = new byte[] { 1 }, Type = ImageType.Emf });

            byte[] package;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                package = ms.ToArray();
            }

            using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(package), System.IO.Compression.ZipArchiveMode.Read);
            var contentTypes = new StreamReader(zip.GetEntry("[Content_Types].xml").Open()).ReadToEnd();
            Assert.Contains("Extension=\"emf\"", contentTypes);
            Assert.Contains("ContentType=\"image/x-emf\"", contentTypes);
        }

        [Fact]
        public void MtefReader_UsesOfficeMathNamespace()
        {
            var mtef = new byte[]
            {
                0x03, 0x01, 0x01, 0x03, 0x00,
                0x02,
                0x78, 0x00,
                0x00
            };

            var omml = new Nedev.FileConverters.DocToDocx.Readers.MtefReader(mtef).ConvertToOmml();

            Assert.NotNull(omml);
            Assert.Contains("http://schemas.openxmlformats.org/officeDocument/2006/math", omml);
            Assert.DoesNotContain("http://schemas.openxmlformats.org/wordprocessingml/2006/main", omml);
        }

        [Fact]
        public void CommentsOnEmptyParagraph_AreMappedCorrectly()
        {
            // include a quick sanity check that the public SaveDocument helper honors
            // the hyperlink toggle (this is largely exercising the convenience API).
            var tmp = Path.GetTempFileName() + ".docx";
            var sample = new DocumentModel();
            sample.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "a", IsHyperlink = true, HyperlinkUrl = "http://x" } } });
            DocToDocxConverter.SaveDocument(sample, tmp, enableHyperlinks: false);
            using (var archive = new System.IO.Compression.ZipArchive(File.OpenRead(tmp), System.IO.Compression.ZipArchiveMode.Read))
            {
                var rels = new StreamReader(archive.GetEntry("word/_rels/document.xml.rels").Open()).ReadToEnd();
                Assert.DoesNotContain("hyperlink", rels, StringComparison.OrdinalIgnoreCase);
            }
            File.Delete(tmp);

            var doc = new DocumentModel();
            var para1 = new ParagraphModel();
            para1.Runs.Add(new RunModel { Text = "hello" });
            var para2 = new ParagraphModel(); // empty
            doc.Paragraphs.Add(para1);
            doc.Paragraphs.Add(para2);
            doc.Annotations.Add(new AnnotationModel { Id = "1", Author = "x", StartCharacterPosition = 0, EndCharacterPosition = 0 });
            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(doc);
                zw.Dispose();
                pkg = ms.ToArray();
            }
            var xml = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read).GetEntry("word/document.xml").Open()).ReadToEnd();
            // comment range start should be inside first paragraph, not lost
            Assert.Contains("commentRangeStart", xml);
        }

        [Fact]
        public void GeneratedDoc_IsValidOpenXml()
        {
            var model = new DocumentModel();
            model.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Check" } } });
            string tmp = Path.GetTempFileName() + ".docx";
            using (var fs = File.Create(tmp))
            {
                using var zw = new ZipWriter(fs);
                zw.WriteDocument(model);
                zw.Dispose();
            }
            using (var package = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(tmp, false))
            {
                var validator = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
                var errors = validator.Validate(package).ToList();
                if (errors.Count > 0)
                {
                    // dump detailed information to help debugging
                    foreach (var err in errors)
                    {
                        _output.WriteLine("Validation error: " + err.Description);
                        _output.WriteLine("  Path: " + err.Path);
                        _output.WriteLine("  Part: " + err.Part.Uri);
                        if (err.Node != null)
                        {
                            _output.WriteLine("  Node XML: " + err.Node.OuterXml);
                        }
                    }
                }
                Assert.Empty(errors);
            }
            File.Delete(tmp);
        }

        [Fact]
        public void SampleTextDoc_Conversion_PreservesRecoveredRunStyling()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "text.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var stylesXml = new StreamReader(archive.GetEntry("word/styles.xml").Open()).ReadToEnd();
                var numberingXml = new StreamReader(archive.GetEntry("word/numbering.xml").Open()).ReadToEnd();
                var visibleText = Regex.Replace(documentXml, "<[^>]+>", string.Empty);
                var xDocument = XDocument.Parse(documentXml);
                var stylesDocument = XDocument.Parse(stylesXml);
                var numberingDocument = XDocument.Parse(numberingXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                Assert.Contains("粗体斜体下划线中横线颜色背景色字体大小字体上标下标组合", visibleText);
                Assert.Contains("<w:pStyle", documentXml);
                Assert.Contains("<w:pStyle w:val=\"Normal\"", documentXml);
                Assert.Contains("<w:b", documentXml);
                Assert.Contains("<w:i", documentXml);
                Assert.Contains("<w:u w:val=\"single\" />", documentXml);
                Assert.Contains("<w:color w:val=\"FF0000\" />", documentXml);
                Assert.Contains("<w:highlight w:val=\"yellow\" />", documentXml);
                Assert.Contains("<w:sz w:val=\"44\" />", documentXml);
                Assert.Contains("<w:vertAlign w:val=\"superscript\" />", documentXml);
                Assert.Contains("<w:vertAlign w:val=\"subscript\" />", documentXml);
                Assert.Contains("<w:bdr ", documentXml);
                Assert.Contains("<w:eastAsianLayout ", documentXml);
                Assert.DoesNotContain("<w:shadow", documentXml, StringComparison.Ordinal);
                Assert.DoesNotContain("<w:del", documentXml, StringComparison.Ordinal);
                Assert.DoesNotContain("<w:ins", documentXml, StringComparison.Ordinal);
                Assert.Contains("styleId=\"Normal\"", stylesXml);
                Assert.Equal(w.NamespaceName, numberingDocument.Root?.Name.NamespaceName);

                var centeredParagraph = FindParagraphContainingText(xDocument, w, "居中");
                Assert.Equal("center", centeredParagraph.Element(w + "pPr")?.Element(w + "jc")?.Attribute(w + "val")?.Value);
                var centeredFonts = centeredParagraph
                    .Descendants(w + "rFonts")
                    .Select(fonts => new
                    {
                        Ascii = fonts.Attribute(w + "ascii")?.Value,
                        EastAsia = fonts.Attribute(w + "eastAsia")?.Value,
                        HAnsi = fonts.Attribute(w + "hAnsi")?.Value,
                        Cs = fonts.Attribute(w + "cs")?.Value
                    })
                    .ToList();
                Assert.Contains(centeredFonts, fonts =>
                    string.Equals(fonts.EastAsia, "SimSun", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(fonts.Ascii, "SimSun", StringComparison.OrdinalIgnoreCase));

                var rightAlignedParagraph = FindParagraphContainingText(xDocument, w, "右对齐");
                Assert.Equal("right", rightAlignedParagraph.Element(w + "pPr")?.Element(w + "jc")?.Attribute(w + "val")?.Value);
                var rightSpacing = rightAlignedParagraph.Element(w + "pPr")?.Element(w + "spacing");
                Assert.NotNull(rightSpacing);
                Assert.Equal("278", rightSpacing!.Attribute(w + "line")?.Value);
                Assert.Equal("auto", rightSpacing.Attribute(w + "lineRule")?.Value);

                var rightAfter = rightSpacing.Attribute(w + "after")?.Value;
                var rightAfterLines = rightSpacing.Attribute(w + "afterLines")?.Value;
                var preservesExpectedAfterSpacing = string.Equals(rightAfterLines, "51", StringComparison.Ordinal) ||
                    string.Equals(rightAfter, "160", StringComparison.Ordinal);
                Assert.True(preservesExpectedAfterSpacing,
                    $"Expected the sample right-aligned paragraph to preserve about 0.51 lines of spacing below, but got after='{rightAfter ?? "<null>"}', afterLines='{rightAfterLines ?? "<null>"}'.");

                var docGrid = xDocument.Descendants(w + "docGrid").FirstOrDefault();
                Assert.NotNull(docGrid);
                Assert.Equal("312", docGrid!.Attribute(w + "linePitch")?.Value);

                var indentParagraph = FindParagraphContainingText(xDocument, w, "Indent");
                Assert.NotNull(indentParagraph.Element(w + "pPr")?.Element(w + "ind"));

                var lineSpacingParagraph = FindParagraphContainingText(xDocument, w, "行间距");
                Assert.NotNull(lineSpacingParagraph.Element(w + "pPr")?.Element(w + "spacing"));

                var scalingParagraph = FindParagraphContainingText(xDocument, w, "文字Scaling 200%");
                Assert.Equal("200", scalingParagraph
                    .Descendants(w + "rPr")
                    .Elements(w + "w")
                    .Attributes(w + "val")
                    .Select(attribute => attribute.Value)
                    .FirstOrDefault());

                var verticalParagraph = FindParagraphContainingText(xDocument, w, "纵向");
                Assert.True(verticalParagraph.Descendants(w + "eastAsianLayout").Any(), "Expected the sample vertical text run to keep eastAsianLayout metadata.");

                var referencedParagraphStyleIds = xDocument
                    .Descendants(w + "pStyle")
                    .Attributes(w + "val")
                    .Select(attribute => attribute.Value)
                    .Distinct(StringComparer.Ordinal)
                    .ToList();

                var definedParagraphStyleIds = stylesDocument
                    .Descendants(w + "style")
                    .Where(element => string.Equals(element.Attribute(w + "type")?.Value, "paragraph", StringComparison.Ordinal))
                    .Attributes(w + "styleId")
                    .Select(attribute => attribute.Value)
                    .ToHashSet(StringComparer.Ordinal);

                Assert.All(referencedParagraphStyleIds, styleId =>
                    Assert.Contains(styleId, definedParagraphStyleIds));

                var usedNumIds = xDocument
                    .Descendants(w + "numPr")
                    .Elements(w + "numId")
                    .Attributes(w + "val")
                    .Select(attribute => attribute.Value)
                    .Distinct(StringComparer.Ordinal)
                    .ToList();

                Assert.NotEmpty(usedNumIds);

                var definedNumIds = numberingDocument
                    .Descendants(w + "num")
                    .Attributes(w + "numId")
                    .Select(attribute => attribute.Value)
                    .ToHashSet(StringComparer.Ordinal);

                Assert.All(usedNumIds, numId => Assert.Contains(numId, definedNumIds));
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTextDoc_LoadDocument_PreservesParagraphLayoutProperties()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);

            var centeredParagraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains("居中", StringComparison.Ordinal));
            Assert.NotNull(centeredParagraph);
            Assert.Equal(ParagraphAlignment.Center, centeredParagraph!.Properties?.Alignment);
            var centeredRunFonts = centeredParagraph.Runs
                .Select(run => run.Properties?.FontName)
                .Where(fontName => !string.IsNullOrWhiteSpace(fontName))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            Assert.Contains("SimSun", centeredRunFonts, StringComparer.OrdinalIgnoreCase);

            var rightAlignedParagraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains("右对齐", StringComparison.Ordinal));
            Assert.NotNull(rightAlignedParagraph);
            Assert.Equal(ParagraphAlignment.Right, rightAlignedParagraph!.Properties?.Alignment);
            Assert.Equal(278, rightAlignedParagraph.Properties?.LineSpacing);
            Assert.Equal(1, rightAlignedParagraph.Properties?.LineSpacingMultiple);
            Assert.True(
                rightAlignedParagraph.Properties?.SpaceAfterLines == 51 || rightAlignedParagraph.Properties?.SpaceAfter == 160,
                $"Expected the right-aligned sample paragraph to preserve about 0.51 lines of spacing below, but got after={rightAlignedParagraph.Properties?.SpaceAfter}, afterLines={rightAlignedParagraph.Properties?.SpaceAfterLines}.");
            Assert.Contains(document.Properties.Sections, section => section.DocGridLinePitch == 312);

            var indentParagraph = document.Paragraphs.FirstOrDefault(p => string.Equals(p.Text, "Indent", StringComparison.Ordinal));
            Assert.NotNull(indentParagraph);
            Assert.True(indentParagraph!.Properties?.IndentLeft > 0, "Expected sample indent paragraph to retain a positive left indent.");
            var lineSpacingParagraph = document.Paragraphs.FirstOrDefault(p => p.Text.Contains("行间距", StringComparison.Ordinal));
            Assert.NotNull(lineSpacingParagraph);
            Assert.True(
                (lineSpacingParagraph!.Properties?.LineSpacing ?? 240) != 240 ||
                (lineSpacingParagraph.Properties?.LineSpacingMultiple ?? 1) != 1,
                "Expected sample line-spacing paragraph to keep a non-default line spacing setting.");

            var listParagraphs = document.Paragraphs
                .Where(p => p.Text.Length == 1 && p.Text[0] is 'A' or 'B' or 'C' or 'D')
                .Take(4)
                .ToList();

            Assert.Equal(4, listParagraphs.Count);
            Assert.All(listParagraphs, paragraph => Assert.True(paragraph.ListFormatId > 0, $"Expected list paragraph '{paragraph.Text}' to retain numbering metadata."));
            Assert.Contains(listParagraphs, paragraph => paragraph.ListLevel > 0);
        }

        [Fact]
        public void SampleTextDoc_NormalStyle_PreservesParagraphSpacingDefaults()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "text.doc");

            using var docReader = new DocReader(inputPath);
            docReader.Load();

            var normalStyle = docReader.Document.Styles.Styles
                .Where(s => s.Type == StyleType.Paragraph)
                .FirstOrDefault(s => string.Equals(s.Name, "Normal", StringComparison.OrdinalIgnoreCase) || s.StyleId == 0 || s.StyleId == StyleIds.NORMAL);

            var details = $"styles={string.Join(" | ", docReader.Document.Styles.Styles.Where(s => s.Type == StyleType.Paragraph).Select(s => $"id={s.StyleId},name={s.Name},line={s.ParagraphProperties?.LineSpacing},mult={s.ParagraphProperties?.LineSpacingMultiple},after={s.ParagraphProperties?.SpaceAfter},afterLines={s.ParagraphProperties?.SpaceAfterLines}"))}";

            Assert.NotNull(normalStyle);
            Assert.NotNull(normalStyle!.ParagraphProperties);
            Assert.True(normalStyle.ParagraphProperties!.LineSpacing == 278, details);
            Assert.True(normalStyle.ParagraphProperties.SpaceAfterLines == 51 || normalStyle.ParagraphProperties.SpaceAfter == 160, details);
        }

        [Fact]
        public void SampleList1Doc_LoadDocument_PreservesParagraphAListMetadata()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "list1.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);

            var paragraphA = document.Paragraphs.FirstOrDefault(p => string.Equals(p.Text, "A", StringComparison.Ordinal));
            Assert.NotNull(paragraphA);
            Assert.True(
                (paragraphA!.Properties?.ListFormatId ?? paragraphA.ListFormatId) > 0,
                "Expected paragraph 'A' to retain list metadata from the source DOC.");

            var listInstanceId = paragraphA.Properties?.ListFormatId ?? paragraphA.ListFormatId;
            var listOverride = document.ListFormatOverrides
                .FirstOrDefault(overrideDefinition => overrideDefinition.OverrideId == listInstanceId);

            Assert.NotNull(listOverride);

            var listDefinition = document.NumberingDefinitions
                .FirstOrDefault(definition => definition.Id == (listOverride?.ListId ?? listInstanceId));

            Assert.NotNull(listDefinition);

            var usedLevels = listDefinition!.Levels.Take(5).ToList();
            Assert.Equal(5, usedLevels.Count);
            Assert.Equal(usedLevels[0].Text, usedLevels[3].Text);
            Assert.Equal(usedLevels[1].Text, usedLevels[4].Text);
            Assert.NotEqual(usedLevels[0].Text, usedLevels[1].Text);
            Assert.NotEqual(usedLevels[1].Text, usedLevels[2].Text);
            Assert.NotEqual(usedLevels[0].RunProperties?.FontName, usedLevels[1].RunProperties?.FontName);
            Assert.NotEqual(usedLevels[1].RunProperties?.FontName, usedLevels[2].RunProperties?.FontName);
        }

        [Fact]
        public void SampleList1Doc_Conversion_PreservesMultilevelBulletGlyphPattern()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "list1.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentEntry = archive.GetEntry("word/document.xml");
                var numberingEntry = archive.GetEntry("word/numbering.xml");
                Assert.NotNull(documentEntry);
                Assert.NotNull(numberingEntry);

                var documentXml = XDocument.Load(documentEntry!.Open());
                var numberingXml = XDocument.Load(numberingEntry!.Open());
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                string[] paragraphTexts = ["A", "c", "d", "e", "f"];
                var levelDetails = paragraphTexts
                    .Select(text => ResolveListLevel(numberingXml, documentXml, w, text))
                    .ToList();

                Assert.All(levelDetails, detail => Assert.False(string.IsNullOrWhiteSpace(detail.LevelText)));
                Assert.Equal(levelDetails[0].LevelText, levelDetails[3].LevelText);
                Assert.Equal(levelDetails[1].LevelText, levelDetails[4].LevelText);
                Assert.NotEqual(levelDetails[0].LevelText, levelDetails[1].LevelText);
                Assert.NotEqual(levelDetails[1].LevelText, levelDetails[2].LevelText);
                Assert.Equal(levelDetails[0].LevelFont, levelDetails[3].LevelFont);
                Assert.Equal(levelDetails[1].LevelFont, levelDetails[4].LevelFont);
                Assert.NotEqual(levelDetails[0].LevelFont, levelDetails[1].LevelFont);
                Assert.NotEqual(levelDetails[1].LevelFont, levelDetails[2].LevelFont);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }

            static (string LevelText, string? LevelFont) ResolveListLevel(XDocument numberingXml, XDocument documentXml, XNamespace w, string paragraphText)
            {
                var paragraph = documentXml
                    .Descendants(w + "p")
                    .FirstOrDefault(element => string.Equals(string.Concat(element.Descendants(w + "t").Select(text => text.Value)), paragraphText, StringComparison.Ordinal));

                Assert.NotNull(paragraph);

                var numPr = paragraph!.Element(w + "pPr")?.Element(w + "numPr");
                Assert.NotNull(numPr);

                var numId = numPr!.Element(w + "numId")?.Attribute(w + "val")?.Value;
                var ilvl = numPr.Element(w + "ilvl")?.Attribute(w + "val")?.Value;
                Assert.False(string.IsNullOrWhiteSpace(numId));
                Assert.False(string.IsNullOrWhiteSpace(ilvl));

                var num = numberingXml
                    .Descendants(w + "num")
                    .FirstOrDefault(element => string.Equals(element.Attribute(w + "numId")?.Value, numId, StringComparison.Ordinal));
                Assert.NotNull(num);

                var abstractNumId = num!.Element(w + "abstractNumId")?.Attribute(w + "val")?.Value;
                Assert.False(string.IsNullOrWhiteSpace(abstractNumId));

                var level = numberingXml
                    .Descendants(w + "abstractNum")
                    .FirstOrDefault(element => string.Equals(element.Attribute(w + "abstractNumId")?.Value, abstractNumId, StringComparison.Ordinal))
                    ?.Elements(w + "lvl")
                    .FirstOrDefault(element => string.Equals(element.Attribute(w + "ilvl")?.Value, ilvl, StringComparison.Ordinal));

                Assert.NotNull(level);

                return (
                    level!.Element(w + "lvlText")?.Attribute(w + "val")?.Value ?? string.Empty,
                    level.Element(w + "rPr")?.Element(w + "rFonts")?.Attribute(w + "ascii")?.Value);
            }
        }

        [Fact]
        public void SampleTableDoc_Conversion_DoesNotHang()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var loadTask = Task.Run(() => DocToDocxConverter.LoadDocument(inputPath));
            Assert.True(loadTask.Wait(TimeSpan.FromSeconds(5)), "Loading sample table.doc took too long (possible hang)");
            Assert.NotNull(loadTask.Result);
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesFirstTableDimensions()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var firstTable = Assert.Single(document.Tables.Where(table => table.ParentTableIndex == null && table.StartParagraphIndex == 0));

            Assert.Equal(4, firstTable.RowCount);
            Assert.Equal(3, firstTable.ColumnCount);
            Assert.All(firstTable.Rows, row => Assert.Equal(3, row.Cells.Count));
            Assert.Equal("a", firstTable.Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Equal("b", firstTable.Rows[0].Cells[1].Paragraphs[0].Text);
            Assert.Equal("c", firstTable.Rows[0].Cells[2].Paragraphs[0].Text);
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesFirstTableFirstColumnAlignment()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var firstTable = document.Tables.Where(table => table.ParentTableIndex == null).OrderBy(table => table.StartParagraphIndex).First();

            Assert.Equal("D", firstTable.Rows[1].Cells[0].Paragraphs[0].Text);
            Assert.Equal(ParagraphAlignment.Center, firstTable.Rows[1].Cells[0].Paragraphs[0].Properties?.Alignment);
            Assert.Equal("中文1", firstTable.Rows[2].Cells[0].Paragraphs[0].Text);
            Assert.Equal(ParagraphAlignment.Right, firstTable.Rows[2].Cells[0].Paragraphs[0].Properties?.Alignment);
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesFirstTableDimensions()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var firstTable = xDocument
                    .Descendants(w + "tbl")
                    .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                    .First();

                var rows = firstTable.Elements(w + "tr").ToList();
                var gridColumns = firstTable.Element(w + "tblGrid")?.Elements(w + "gridCol").ToList();

                Assert.Equal(4, rows.Count);
                Assert.NotNull(gridColumns);
                Assert.Equal(3, gridColumns!.Count);
                Assert.All(rows, row => Assert.Equal(3, row.Elements(w + "tc").Count()));

                var firstRowTexts = rows[0]
                    .Elements(w + "tc")
                    .Select(cell => string.Concat(cell.Descendants(w + "t").Select(text => text.Value)))
                    .ToList();

                Assert.Equal(new[] { "a", "b", "c" }, firstRowTexts);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesFirstTableEqualColumnWidths()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var firstTable = document.Tables
                .Where(table => table.ParentTableIndex == null)
                .OrderBy(table => table.StartParagraphIndex)
                .First();

            Assert.Equal(3, firstTable.ColumnCount);
            Assert.True(firstTable.Rows.Count > 0, "Expected the first top-level table to contain at least one row.");
            Assert.Equal(3, firstTable.Rows[0].Cells.Count);
            AssertTwipsClose(CmToTwips(5.43), firstTable.Rows[0].Cells[0].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(5.43), firstTable.Rows[0].Cells[1].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(5.43), firstTable.Rows[0].Cells[2].Properties?.Width ?? 0);
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesFirstTableEqualColumnWidths()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var firstTable = xDocument
                    .Descendants(w + "tbl")
                    .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                    .First();

                var gridColumns = firstTable.Element(w + "tblGrid")?.Elements(w + "gridCol").ToList();

                Assert.NotNull(gridColumns);
                Assert.Equal(3, gridColumns!.Count);
                AssertTwipsClose(CmToTwips(5.43), (int?)gridColumns[0].Attribute(w + "w") ?? 0);
                AssertTwipsClose(CmToTwips(5.43), (int?)gridColumns[1].Attribute(w + "w") ?? 0);
                AssertTwipsClose(CmToTwips(5.43), (int?)gridColumns[2].Attribute(w + "w") ?? 0);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesZeroSectionGutter()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);

            Assert.NotEmpty(document.Properties.Sections);
            Assert.All(document.Properties.Sections, section => Assert.Equal(0, section.Gutter));
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesZeroSectionGutter()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var margins = xDocument
                    .Descendants(w + "sectPr")
                    .Elements(w + "pgMar")
                    .ToList();

                Assert.NotEmpty(margins);
                Assert.All(margins, margin => Assert.Equal(0, (int?)margin.Attribute(w + "gutter") ?? -1));
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesBorderlessTableAfterMarkerParagraph()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var markerParagraph = document.Paragraphs.FirstOrDefault(paragraph =>
                paragraph.Text.Contains("无边框", StringComparison.Ordinal));

            Assert.NotNull(markerParagraph);

            var targetTable = document.Tables
                .Where(table => table.ParentTableIndex == null && table.StartParagraphIndex > markerParagraph!.Index)
                .OrderBy(table => table.StartParagraphIndex)
                .FirstOrDefault();

            Assert.NotNull(targetTable);
            Assert.True(targetTable!.Properties == null ||
                        (targetTable.Properties.BorderTop == null &&
                         targetTable.Properties.BorderBottom == null &&
                         targetTable.Properties.BorderLeft == null &&
                         targetTable.Properties.BorderRight == null &&
                         targetTable.Properties.BorderInsideH == null &&
                         targetTable.Properties.BorderInsideV == null),
                "Expected the top-level table after the 无边框 paragraph to remain borderless in the parsed model.");
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesBorderlessTableAfterMarkerParagraph()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var bodyElements = xDocument.Root?
                    .Element(w + "body")?
                    .Elements()
                    .ToList();

                Assert.NotNull(bodyElements);

                var markerIndex = bodyElements!.FindIndex(element =>
                    element.Name == w + "p" &&
                    string.Concat(element.Descendants(w + "t").Select(text => text.Value)).Contains("无边框", StringComparison.Ordinal));

                Assert.True(markerIndex >= 0, "Expected converted document to contain the 无边框 paragraph.");

                var targetTable = bodyElements
                    .Skip(markerIndex + 1)
                    .FirstOrDefault(element => element.Name == w + "tbl");

                Assert.NotNull(targetTable);
                Assert.Null(targetTable!.Element(w + "tblPr")?.Element(w + "tblBorders"));
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesFirstTableFirstColumnAlignment()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var firstTable = xDocument
                    .Descendants(w + "tbl")
                    .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                    .First();
                var rows = firstTable.Elements(w + "tr").ToList();

                Assert.True(rows.Count >= 3, "Expected the first top-level table to have at least three rows.");

                var secondRowFirstCellParagraph = rows[1].Elements(w + "tc").First().Elements(w + "p").First();
                var thirdRowFirstCellParagraph = rows[2].Elements(w + "tc").First().Elements(w + "p").First();

                var secondRowText = string.Concat(secondRowFirstCellParagraph.Descendants(w + "t").Select(text => text.Value));
                var thirdRowText = string.Concat(thirdRowFirstCellParagraph.Descendants(w + "t").Select(text => text.Value));
                var secondRowAlignment = secondRowFirstCellParagraph.Element(w + "pPr")?.Element(w + "jc")?.Attribute(w + "val")?.Value ?? "(none)";
                var thirdRowAlignment = thirdRowFirstCellParagraph.Element(w + "pPr")?.Element(w + "jc")?.Attribute(w + "val")?.Value ?? "(none)";

                Assert.Equal("D", secondRowText);
                Assert.Equal("中文1", thirdRowText);
                Assert.Equal("center", secondRowAlignment);
                Assert.Equal("right", thirdRowAlignment);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesThirdTopLevelTableFixedColumnWidths()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var markerParagraph = document.Paragraphs.FirstOrDefault(paragraph =>
                paragraph.Text.Contains("固定列宽", StringComparison.Ordinal));

            Assert.NotNull(markerParagraph);

            var thirdTable = document.Tables
                .Where(table => table.ParentTableIndex == null && table.StartParagraphIndex > markerParagraph!.Index)
                .OrderBy(table => table.StartParagraphIndex)
                .First(table =>
                    table.Rows.Count > 0 &&
                    table.Rows[0].Cells.Count >= 2 &&
                    string.Equals(
                        string.Concat(table.Rows[0].Cells[0].Paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text)).Trim(),
                        "固定列宽",
                        StringComparison.Ordinal));

            Assert.Equal(3, thirdTable.ColumnCount);
            Assert.True(thirdTable.Rows.Count > 0, "Expected the fixed-width table to contain at least one row.");
            Assert.Equal(3, thirdTable.Rows[0].Cells.Count);
            AssertTwipsClose(CmToTwips(2.19), thirdTable.Rows[0].Cells[0].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(11.25), thirdTable.Rows[0].Cells[1].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(2.86), thirdTable.Rows[0].Cells[2].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(16.30), thirdTable.Properties?.PreferredWidth ?? 0);
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesFixedColumnWidthsForTableAfterMarkerParagraph()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var markerParagraph = document.Paragraphs.FirstOrDefault(paragraph =>
                paragraph.Text.Contains("固定列宽", StringComparison.Ordinal));

            Assert.NotNull(markerParagraph);

            var targetTable = document.Tables
                .Where(table => table.ParentTableIndex == null && table.StartParagraphIndex > markerParagraph!.Index)
                .OrderBy(table => table.StartParagraphIndex)
                .FirstOrDefault(table =>
                    table.Rows.Count > 0 &&
                    table.Rows[0].Cells.Count == 3 &&
                    string.Equals(
                        string.Concat(table.Rows[0].Cells[0].Paragraphs.SelectMany(paragraph => paragraph.Runs).Select(run => run.Text)).Trim(),
                        "固定列宽",
                        StringComparison.Ordinal));

            Assert.NotNull(targetTable);
            Assert.Equal(3, targetTable!.ColumnCount);
            AssertTwipsClose(CmToTwips(2.19), targetTable.Rows[0].Cells[0].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(11.25), targetTable.Rows[0].Cells[1].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(2.86), targetTable.Rows[0].Cells[2].Properties?.Width ?? 0);
            AssertTwipsClose(CmToTwips(16.30), targetTable.Properties?.PreferredWidth ?? 0);
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesThirdTopLevelTableFixedColumnWidths()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var thirdTable = xDocument
                    .Descendants(w + "tbl")
                    .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                    .First(tbl =>
                    {
                        var firstCellParagraph = tbl
                            .Elements(w + "tr")
                            .FirstOrDefault()?
                            .Elements(w + "tc")
                            .FirstOrDefault()?
                            .Elements(w + "p")
                            .FirstOrDefault();

                        var firstCellText = firstCellParagraph == null
                            ? string.Empty
                            : string.Concat(firstCellParagraph.Descendants(w + "t").Select(text => text.Value)).Trim();

                        return string.Equals(firstCellText, "固定列宽", StringComparison.Ordinal);
                    });

                var gridColumns = thirdTable.Elements(w + "tblGrid").Elements(w + "gridCol").ToList();

                Assert.Equal(3, gridColumns.Count);
                AssertTwipsClose(CmToTwips(2.19), (int?)gridColumns[0].Attribute(w + "w") ?? 0);
                AssertTwipsClose(CmToTwips(11.25), (int?)gridColumns[1].Attribute(w + "w") ?? 0);
                AssertTwipsClose(CmToTwips(2.86), (int?)gridColumns[2].Attribute(w + "w") ?? 0);
                AssertTwipsClose(CmToTwips(16.30), (int?)thirdTable.Element(w + "tblPr")?.Element(w + "tblW")?.Attribute(w + "w") ?? 0);
                Assert.Equal("fixed", (string?)thirdTable.Element(w + "tblPr")?.Element(w + "tblLayout")?.Attribute(w + "type"));
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesThirdTopLevelTableVisibleBorders()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var thirdTable = xDocument
                    .Descendants(w + "tbl")
                    .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                    .Skip(2)
                    .First();

                Assert.NotNull(thirdTable.Element(w + "tblPr")?.Element(w + "tblBorders"));
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesBorderColorTableAfterMarkerParagraph()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var markerParagraph = document.Paragraphs.FirstOrDefault(paragraph =>
                paragraph.Text.Contains("边框颜色", StringComparison.Ordinal));

            Assert.NotNull(markerParagraph);

            var targetTable = document.Tables
                .Where(table => table.ParentTableIndex == null && table.StartParagraphIndex > markerParagraph!.Index)
                .OrderBy(table => table.StartParagraphIndex)
                .FirstOrDefault();

            Assert.NotNull(targetTable);

            var borders = new[]
            {
                targetTable!.Properties?.BorderTop,
                targetTable.Properties?.BorderBottom,
                targetTable.Properties?.BorderLeft,
                targetTable.Properties?.BorderRight,
                targetTable.Properties?.BorderInsideH,
                targetTable.Properties?.BorderInsideV
            }
            .Concat(targetTable.Rows
                .SelectMany(row => row.Cells)
                .SelectMany(cell => new[]
                {
                    cell.Properties?.BorderTop,
                    cell.Properties?.BorderBottom,
                    cell.Properties?.BorderLeft,
                    cell.Properties?.BorderRight
                }))
            .ToArray();

            Assert.Contains(borders, border => border != null && border.Color > 16);
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesBorderColorTableAfterMarkerParagraph()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var bodyElements = xDocument.Root?
                    .Element(w + "body")?
                    .Elements()
                    .ToList();

                Assert.NotNull(bodyElements);

                var markerIndex = bodyElements!.FindIndex(element =>
                    element.Name == w + "p" &&
                    string.Concat(element.Descendants(w + "t").Select(text => text.Value)).Contains("边框颜色", StringComparison.Ordinal));

                Assert.True(markerIndex >= 0, "Expected converted document to contain the 边框颜色 paragraph.");

                var targetTable = bodyElements
                    .Skip(markerIndex + 1)
                    .FirstOrDefault(element => element.Name == w + "tbl");

                Assert.NotNull(targetTable);

                var borders = targetTable!
                    .Element(w + "tblPr")?
                    .Element(w + "tblBorders")?
                    .Elements()
                    .ToList();

                Assert.NotNull(borders);

                var loadedDocument = DocToDocxConverter.LoadDocument(inputPath);
                var loadedMarkerParagraph = loadedDocument.Paragraphs.First(paragraph =>
                    paragraph.Text.Contains("边框颜色", StringComparison.Ordinal));
                var loadedTable = loadedDocument.Tables
                    .Where(table => table.ParentTableIndex == null && table.StartParagraphIndex > loadedMarkerParagraph.Index)
                    .OrderBy(table => table.StartParagraphIndex)
                    .First();

                var expectedColors = new[]
                {
                    loadedTable.Properties?.BorderTop?.Color,
                    loadedTable.Properties?.BorderBottom?.Color,
                    loadedTable.Properties?.BorderLeft?.Color,
                    loadedTable.Properties?.BorderRight?.Color,
                    loadedTable.Properties?.BorderInsideH?.Color,
                    loadedTable.Properties?.BorderInsideV?.Color
                }
                .Concat(loadedTable.Rows
                    .SelectMany(row => row.Cells)
                    .SelectMany(cell => new int?[]
                    {
                        cell.Properties?.BorderTop?.Color,
                        cell.Properties?.BorderBottom?.Color,
                        cell.Properties?.BorderLeft?.Color,
                        cell.Properties?.BorderRight?.Color
                    }))
                .Where(color => color.HasValue && color.Value > 16)
                .Select(color => ColorHelper.ResolveColorHex(color!.Value, loadedDocument.Theme))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

                Assert.NotEmpty(expectedColors);
                Assert.DoesNotContain(borders!, border =>
                    string.Equals(border.Attribute(w + "color")?.Value, "D3D3D3", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(borders!, border =>
                    expectedColors.Contains(border.Attribute(w + "color")?.Value ?? string.Empty, StringComparer.OrdinalIgnoreCase));
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_LoadDocument_PreservesFourthTopLevelTableFixedRowHeightColumns()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");

            var document = DocToDocxConverter.LoadDocument(inputPath);
            var topLevelTables = document.Tables
                .Where(table => table.ParentTableIndex == null)
                .OrderBy(table => table.StartParagraphIndex)
                .ToList();
            var fixedRowHeightTable = topLevelTables.Skip(3).First();

            Assert.Equal(2, fixedRowHeightTable.RowCount);
            Assert.Equal(3, fixedRowHeightTable.ColumnCount);
            Assert.All(fixedRowHeightTable.Rows, row => Assert.Equal(3, row.Cells.Count));
            Assert.Equal("Afdasfdsfdsfvscer测试自动换行", fixedRowHeightTable.Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Equal("固定行高", fixedRowHeightTable.Rows[1].Cells[0].Paragraphs[0].Text);
            Assert.NotNull(fixedRowHeightTable.Rows[1].Properties);
            AssertTwipsClose(CmToTwips(2.79), fixedRowHeightTable.Rows[1].Properties!.Height);
            Assert.True(fixedRowHeightTable.Rows[1].Properties.HeightIsExact);
            Assert.All(fixedRowHeightTable.Rows.SelectMany(row => row.Cells.Skip(1)), cell =>
                Assert.True(string.IsNullOrWhiteSpace(cell.Paragraphs[0].Text)));
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesFourthTopLevelTableFixedRowHeightColumns()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var fixedRowHeightTable = xDocument
                    .Descendants(w + "tbl")
                    .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                    .Skip(3)
                    .First();

                var rows = fixedRowHeightTable.Elements(w + "tr").ToList();
                var gridColumns = fixedRowHeightTable.Elements(w + "tblGrid").Elements(w + "gridCol").ToList();

                Assert.Equal(2, rows.Count);
                Assert.Equal(3, gridColumns.Count);
                Assert.All(rows, row => Assert.Equal(3, row.Elements(w + "tc").Count()));

                var secondRowHeight = rows[1].Element(w + "trPr")?.Element(w + "trHeight");
                Assert.NotNull(secondRowHeight);
                Assert.Equal("exact", secondRowHeight!.Attribute(w + "hRule")?.Value);
                AssertTwipsClose(CmToTwips(2.79), (int?)secondRowHeight.Attribute(w + "val") ?? 0);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesSecondTopLevelTableWithoutMergingFirstRowSecondCell()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                var document = DocToDocxConverter.LoadDocument(inputPath);
                var secondTable = document.Tables
                    .Where(table => table.ParentTableIndex == null)
                    .OrderBy(table => table.StartParagraphIndex)
                    .Skip(1)
                    .First();

                Assert.Equal(2, secondTable.RowCount);
                Assert.Equal(2, secondTable.ColumnCount);
                Assert.Equal(2, secondTable.Rows[0].Cells.Count);
                Assert.Equal(2, secondTable.Rows[1].Cells.Count);
                Assert.Equal(1, secondTable.Rows[0].Cells[0].Paragraphs.Count(p => p.Type == ParagraphType.NestedTable && p.NestedTable != null));
                Assert.Equal(1, secondTable.Rows[0].Cells[1].ColumnSpan);
                Assert.Equal(1, secondTable.Rows[0].Cells[1].RowSpan);

                DocToDocxConverter.SaveDocument(document, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var xmlSecondTable = xDocument
                    .Descendants(w + "tbl")
                    .Where(tbl => !tbl.Ancestors(w + "tbl").Any())
                    .Skip(1)
                    .First();
                var rows = xmlSecondTable.Elements(w + "tr").ToList();

                Assert.Equal(2, rows.Count);

                var firstRowCells = rows[0].Elements(w + "tc").ToList();
                Assert.Equal(2, firstRowCells.Count);
                Assert.True(firstRowCells[0].Descendants(w + "tbl").Any(), "Expected the nested table to stay in the first cell of the second top-level table.");
                Assert.Equal(w + "p", firstRowCells[0].Elements().Last().Name);
                Assert.False(firstRowCells[1].Descendants(w + "gridSpan").Any(), "Expected the first-row second cell to remain a standalone cell without horizontal merge markup.");
                Assert.False(firstRowCells[1].Descendants(w + "vMerge").Any(), "Expected the first-row second cell to remain a standalone cell without vertical merge markup.");
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesNestedTablesAndCenteredHeading()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                var document = DocToDocxConverter.LoadDocument(inputPath);
                Assert.True(CountNestedTables(document.Tables) >= 1, "Expected parsed sample table model to contain nested tables.");

                DocToDocxConverter.SaveDocument(document, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var tables = xDocument.Descendants(w + "tbl").ToList();
                var topLevelTables = tables.Where(tbl => !tbl.Ancestors(w + "tbl").Any()).ToList();

                Assert.True(tables.Count >= 6, "Expected sample table output to contain the recovered tables and one nested table.");
                Assert.True(topLevelTables.Count >= 5, "Expected sample table output to contain five top-level tables.");
                Assert.Contains("表格嵌套", documentXml);
                Assert.True(
                    xDocument.Descendants(w + "tc").Any(tc => tc.Descendants(w + "tbl").Any()),
                    "Expected at least one nested table in the sample output.");

                var nestedTable = xDocument
                    .Descendants(w + "tbl")
                    .FirstOrDefault(tbl => tbl.Ancestors(w + "tbl").Any());

                Assert.NotNull(nestedTable);
                Assert.NotNull(nestedTable!.Element(w + "tblPr")?.Element(w + "tblBorders"));
                Assert.NotNull(nestedTable.Ancestors(w + "tbl").First().Element(w + "tblPr")?.Element(w + "tblBorders"));
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_Conversion_PlacesNestedTableImmediatelyAfterTitle()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var bodyElements = xDocument.Root?
                    .Element(w + "body")?
                    .Elements()
                    .ToList();

                Assert.NotNull(bodyElements);

                var titleIndex = bodyElements!.FindIndex(element =>
                    element.Name == w + "p" &&
                    string.Concat(element.Descendants(w + "t").Select(text => text.Value)).Contains("表格嵌套", StringComparison.Ordinal));

                Assert.True(titleIndex >= 0, "Expected converted document to contain the 表格嵌套 paragraph.");
                Assert.True(titleIndex + 1 < bodyElements.Count, "Expected a body element after the 表格嵌套 paragraph.");
                Assert.Equal(w + "tbl", bodyElements[titleIndex + 1].Name);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleTableDoc_Conversion_PreservesMergedCellSectionAsDirectTwoColumnTable()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "table.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var bodyElements = xDocument.Root?
                    .Element(w + "body")?
                    .Elements()
                    .ToList();

                Assert.NotNull(bodyElements);

                var titleIndex = bodyElements!.FindIndex(element =>
                    element.Name == w + "p" &&
                    string.Concat(element.Descendants(w + "t").Select(text => text.Value)).Contains("单元格合并", StringComparison.Ordinal));

                Assert.True(titleIndex >= 0, "Expected converted document to contain the 单元格合并 paragraph.");
                var mergeTable = bodyElements
                    .Skip(titleIndex + 1)
                    .TakeWhile(element =>
                        element.Name != w + "p" ||
                        string.IsNullOrWhiteSpace(string.Concat(element.Descendants(w + "t").Select(text => text.Value))))
                    .FirstOrDefault(element => element.Name == w + "tbl");

                Assert.NotNull(mergeTable);
                Assert.Equal(w + "tbl", mergeTable.Name);
                Assert.False(mergeTable.Descendants(w + "tbl").Any(), "Expected 单元格合并 to remain a direct top-level table instead of a synthesized nested wrapper.");

                var rows = mergeTable.Elements(w + "tr").ToList();
                Assert.NotEmpty(rows);

                var firstRowCells = rows[0].Elements(w + "tc").ToList();
                Assert.Equal(2, firstRowCells.Count);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleImageDoc_Conversion_PreservesImagePartsAndDrawings()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "image.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                var documentXml = new StreamReader(archive.GetEntry("word/document.xml").Open()).ReadToEnd();
                var relsXml = new StreamReader(archive.GetEntry("word/_rels/document.xml.rels").Open()).ReadToEnd();
                var xDocument = XDocument.Parse(documentXml);
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                var extents = xDocument.Descendants(wp + "extent").ToList();
                var firstExtent = extents.FirstOrDefault();
                var secondExtent = extents.Skip(1).FirstOrDefault();

                Assert.Contains("原始", documentXml);
                Assert.Contains("缩放", documentXml);
                Assert.Contains("对比度", documentXml);
                Assert.Equal(5, Regex.Matches(documentXml, "<w:drawing\\b").Count);
                Assert.Equal(5, Regex.Matches(documentXml, "<a:blip r:embed=").Count);
                Assert.Equal(5, archive.Entries.Count(entry => entry.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase)));
                Assert.Equal(5, Regex.Matches(relsXml, "relationships/image").Count);
                Assert.Equal(5, extents.Count);
                Assert.NotNull(firstExtent);
                Assert.NotNull(secondExtent);
                var firstWidthCm = Math.Round((double)firstExtent!.Attribute("cx")! / 360000d, 2);
                var firstHeightCm = Math.Round((double)firstExtent.Attribute("cy")! / 360000d, 2);
                var secondWidthCm = Math.Round((double)secondExtent!.Attribute("cx")! / 360000d, 2);
                var secondHeightCm = Math.Round((double)secondExtent.Attribute("cy")! / 360000d, 2);
                Assert.InRange(firstWidthCm, 4.02, 4.12);
                Assert.InRange(firstHeightCm, 0.77, 0.87);
                Assert.InRange(secondWidthCm, 15.87, 15.97);
                Assert.InRange(secondHeightCm, 3.33, 3.43);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void SampleImageDoc_Reader_PreservesPngPixelDimensions()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "image.doc");

            using var stream = File.OpenRead(inputPath);
            using var reader = new DocReader(stream, password: null);
            reader.Load();

            var firstImage = reader.Document.Images.FirstOrDefault();

            Assert.NotNull(firstImage);
            Assert.Equal(154, firstImage!.Width);
            Assert.Equal(31, firstImage.Height);
        }

        [Fact]
        public void SampleImageDoc_Conversion_PreservesHorizontalAndVerticalFlip()
        {
            var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", ".."));
            var inputPath = Path.Combine(repoRoot, "samples", "image.doc");
            var outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                DocToDocxConverter.Convert(inputPath, outputPath);

                using var archive = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                using var stream = archive.GetEntry("word/document.xml")!.Open();
                var xDocument = XDocument.Load(stream);
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
                var pictureTransforms = xDocument
                    .Descendants(pic + "spPr")
                    .Elements(a + "xfrm")
                    .ToList();

                Assert.True(pictureTransforms.Count >= 5, $"Expected at least 5 picture transforms, found {pictureTransforms.Count}.");
                Assert.Equal("1", pictureTransforms[3].Attribute("flipH")?.Value);
                Assert.Equal("1", pictureTransforms[4].Attribute("flipV")?.Value);
            }
            finally
            {
                if (File.Exists(outputPath))
                    File.Delete(outputPath);
            }
        }

        [Fact]
        public void WriteDocument_FloatingShape_UsesCustomWrapPolygon()
        {
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Index = 0, Runs = { new RunModel { Text = "body" } } });
            doc.Shapes.Add(new ShapeModel
            {
                Id = 1,
                Type = ShapeType.Rectangle,
                ParagraphIndexHint = 0,
                Anchor = new ShapeAnchor
                {
                    IsFloating = true,
                    Width = 2000,
                    Height = 1200,
                    WrapType = ShapeWrapType.Tight
                },
                WrapPolygonVertices = new List<System.Drawing.Point>
                {
                    new(0, 0),
                    new(10800, 0),
                    new(21600, 10800),
                    new(0, 21600)
                }
            });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                new DocumentWriter(writer).WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<wp:wrapTight", xml);
            Assert.Contains("<wp:start x=\"0\" y=\"0\"", xml);
            Assert.Contains("<wp:lineTo x=\"10800\" y=\"0\"", xml);
            Assert.Contains("<wp:lineTo x=\"21600\" y=\"10800\"", xml);
        }

        [Fact]
        public void WriteDocument_Textbox_UsesWrapAndAlignmentMetadata()
        {
            var doc = new DocumentModel();
            doc.Textboxes.Add(new TextboxModel
            {
                Index = 1,
                Name = "Box 1",
                Left = 100,
                Top = 200,
                Width = 2400,
                Height = 1200,
                WrapMode = TextboxWrapMode.Tight,
                VerticalAlignment = TextboxVerticalAlignment.Center,
                HorizontalAlignment = TextboxHorizontalAlignment.Center,
                Paragraphs =
                {
                    new ParagraphModel { Runs = { new RunModel { Text = "textbox" } } }
                }
            });

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                new DocumentWriter(writer).WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("<wp:wrapTight", xml);
            Assert.Contains("<wps:bodyPr wrap=\"tight\" vert=\"ctr\" anchorCtr=\"1\"", xml);
            Assert.Contains("textbox", xml);
        }

        [Fact]
        public void MergeTextboxShapesIntoTextboxes_FlowsIntoWrittenDocx_WithoutDuplicateTextboxShapes()
        {
            var doc = new DocumentModel();
            doc.Textboxes.Add(new TextboxModel
            {
                Index = 1,
                Paragraphs =
                {
                    new ParagraphModel
                    {
                        Properties = new ParagraphProperties { Alignment = ParagraphAlignment.Center },
                        Runs = { new RunModel { Text = "merged textbox" } }
                    }
                }
            });
            doc.Shapes.Add(new ShapeModel
            {
                Id = 21,
                Type = ShapeType.Textbox,
                Anchor = new ShapeAnchor
                {
                    IsFloating = true,
                    X = 100,
                    Y = 200,
                    Width = 2400,
                    Height = 1200,
                    WrapType = ShapeWrapType.Through
                }
            });

            DocReader.MergeTextboxShapesIntoTextboxes(doc);

            string xml;
            using (var ms = new MemoryStream())
            {
                var settings = new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true };
                using var writer = XmlWriter.Create(ms, settings);
                new DocumentWriter(writer).WriteDocument(doc);
                writer.Flush();
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("merged textbox", xml);
            Assert.Contains("<wp:wrapThrough", xml);
            Assert.Contains("<wps:bodyPr wrap=\"through\"", xml);
            Assert.DoesNotContain("Shape 21", xml);
        }

        private static int CountNestedTables(IEnumerable<TableModel> tables)
        {
            int count = 0;

            foreach (var table in tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var paragraph in cell.Paragraphs)
                        {
                            if (paragraph.Type != ParagraphType.NestedTable || paragraph.NestedTable == null)
                                continue;

                            count++;
                            count += CountNestedTables(new[] { paragraph.NestedTable });
                        }
                    }
                }
            }

            return count;
        }

        private static XElement FindParagraphContainingText(XDocument document, XNamespace w, string text)
        {
            return document
                .Descendants(w + "p")
                .First(paragraph => string.Concat(paragraph.Descendants(w + "t").Select(t => t.Value)).Contains(text, StringComparison.Ordinal));
        }

        private static int CmToTwips(double centimeters)
        {
            return (int)Math.Round((centimeters / 2.54d) * 1440d, MidpointRounding.AwayFromZero);
        }

        private static void AssertTwipsClose(int expected, int actual, int tolerance = 24)
        {
            Assert.InRange(actual, expected - tolerance, expected + tolerance);
        }
    }
}
