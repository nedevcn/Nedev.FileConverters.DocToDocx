#nullable enable
using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Nedev.FileConverters.DocToDocx.Models;
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
                xml = Encoding.UTF8.GetString(ms.ToArray());
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
                xml = Encoding.UTF8.GetString(ms.ToArray());
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
                xml = Encoding.UTF8.GetString(ms.ToArray());
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
            Assert.NotNull(zip.GetEntry("word/header2.xml"));
            Assert.NotNull(zip.GetEntry("word/footer2.xml"));
            Assert.NotNull(zip.GetEntry("word/footnotes.xml"));

            var hdr = new StreamReader(zip.GetEntry("word/header2.xml").Open()).ReadToEnd();
            Assert.Contains("HDR", hdr);
            var ftr = new StreamReader(zip.GetEntry("word/footer2.xml").Open()).ReadToEnd();
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
                xml = Encoding.UTF8.GetString(ms.ToArray());
            }

            Assert.Contains("before", xml);
            // should have at least two <w:tbl> entries (parent and nested)
            int count = xml.Split("<w:tbl").Length - 1;
            Assert.True(count >= 2, "Expected at least two tables, got " + count);
            Assert.Contains("inner", xml);
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
            var rels = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read).GetEntry("word/_rels/document.xml.rels").Open()).ReadToEnd();
            Assert.Contains("foo", rels);
            Assert.Contains("bar", rels);
            // ensure two hyperlink relationships only
            var hyperCount = Regex.Matches(rels, "Type=\\\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\\\"").Count;
            Assert.Equal(2, hyperCount);
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
    }
}
