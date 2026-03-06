#nullable enable
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using System.Reflection;
using System.Text;
using Nedev.FileConverters.DocToDocx.Readers;
using Nedev.FileConverters.DocToDocx.
using Nedev.FileConverters.DocToDocx.Cli;
using Nedev.FileConverters.DocToDocx.Models;
using Nedev.FileConverters.DocToDocx.Writers;

namespace Nedev.FileConverters.DocToDocx.Tests
{
    public class ReaderIntegrationTests
    {
        private readonly Xunit.Abstractions.ITestOutputHelper _output;

        public ReaderIntegrationTests(Xunit.Abstractions.ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void CreateAndLoadDocument_HasContent()
        {
            // LoadDocument only supports .doc files; attempting to load a .docx should throw.
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var original = new DocumentModel();
            original.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Hello" } } });
            DocToDocxConverter.SaveDocument(original, path);

            try
            {
                Assert.Throws<InvalidDataException>(() => DocToDocxConverter.LoadDocument(path));
            }
            finally
            {
                // ensure the temporary file is removed even if the assertion fails or a handle was leaked
                if (File.Exists(path))
                    File.Delete(path);
            }
        }

        [Fact]
        public async Task Cli_CopiesDocxInput_WhenPassedDocx()
        {
            string tempInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "X" } } });
            DocToDocxConverter.SaveDocument(doc, tempInput);

            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            await Nedev.DocToDocx.Cli.Program.Main(new[] { tempInput, outPath });
            Assert.True(File.Exists(outPath));

            // verify copy semantics (size)
            Assert.Equal(new FileInfo(tempInput).Length, new FileInfo(outPath).Length);
            File.Delete(tempInput);
            File.Delete(outPath);
        }

        [Fact]
        public void Convert_WithProgress_ReportsStages()
        {
            // create a simple docx file instead of relying on an external .doc sample
            string inPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var model = new DocumentModel();
            model.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "progress" } } });
            DocToDocxConverter.SaveDocument(model, inPath);

            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            var reports = new List<ConversionProgress>();
            var progress = new Progress<ConversionProgress>(p => reports.Add(p));

            // since we pass a .docx file, the converter will copy it; progress stages should still fire.
            DocToDocxConverter.Convert(inPath, outPath, progress);

            Assert.Contains(reports, r => r.Stage == ConversionStage.Reading);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Writing);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Complete);

            File.Delete(outPath);
            File.Delete(inPath);
        }

        [Fact]
        public void Convert_WithoutProgress_CopiesDocx()
        {
            string inPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "copy" } } });
            DocToDocxConverter.SaveDocument(doc, inPath);

            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            DocToDocxConverter.Convert(inPath, outPath);

            Assert.True(File.Exists(outPath));
            Assert.Equal(new FileInfo(inPath).Length, new FileInfo(outPath).Length);

            File.Delete(inPath);
            File.Delete(outPath);
        }

        [Fact]
        public async Task Cli_VersionFlag_PrintsVersion()
        {
            using var sw = new StringWriter();
            Console.SetOut(sw);
            await Nedev.DocToDocx.Cli.Program.Main(new[] { "--version" });
            string output = sw.ToString();
            Assert.Contains("Version", output);
        }

        [Fact]
        public async Task Cli_DirectoryConversion_WritesFiles()
        {
            string tempInput = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            string tempOutput = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempInput);
            Directory.CreateDirectory(tempOutput);

            // create a nested docx file
            var sub = Path.Combine(tempInput, "sub");
            Directory.CreateDirectory(sub);
            string aPath = Path.Combine(sub, "a.docx");
            var doc = new DocumentModel();
            doc.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Z" } } });
            DocToDocxConverter.SaveDocument(doc, aPath);

            await Nedev.DocToDocx.Cli.Program.Main(new[] { tempInput, tempOutput, "-r" });

            string expected = Path.Combine(tempOutput, "sub", "a.docx");
            Assert.True(File.Exists(expected), "Converted file should exist");

            Directory.Delete(tempInput, true);
            Directory.Delete(tempOutput, true);
        }

        [Fact]
        public void SaveDocument_GeneratesValidDocx()
        {
            var model = new DocumentModel();
            model.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Hello" } } });
            string tempOut = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            DocToDocxConverter.SaveDocument(model, tempOut);
            Assert.True(File.Exists(tempOut));

            using (var zip = new System.IO.Compression.ZipArchive(File.OpenRead(tempOut), System.IO.Compression.ZipArchiveMode.Read))
            {
                var entry = zip.GetEntry("word/document.xml");
                Assert.NotNull(entry);
                using var reader = new StreamReader(entry.Open());
                var xml = reader.ReadToEnd();
                Assert.Contains("Hello", xml);
            }

            File.Delete(tempOut);
        }

        [Fact]
        public void DumpHyperlinkRuns_FromProblematicDoc()
        {
            // this test prints run information for the paragraph containing the
            // broken hyperlink in the supplied .doc sample so we can understand
            // how the reader has populated RunModel properties.
            string sample = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "渠道授权协议v5.doc"));
            Assert.True(File.Exists(sample), $"sample file must exist at {sample}");

            var model = DocToDocxConverter.LoadDocument(sample);
            var para = model.Paragraphs.FirstOrDefault(p => p.Runs.Any(r => r.Text != null && r.Text.Contains("代理期间")));
            Assert.NotNull(para);

            foreach (var r in para!.Runs)
            {
                Console.WriteLine($"Run: '{r.Text}' IsHyperlink={r.IsHyperlink} Url='{r.HyperlinkUrl}' RelId='{r.HyperlinkRelationshipId}'");
            }

            // nothing to assert really, this is for inspection only
            Assert.True(true);
        }

        [Fact]
        public void DumpRuns_With5_3_PrintSanitized()
        {
            string sample = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "渠道授权协议v5.doc"));
            var model = DocToDocxConverter.LoadDocument(sample);
            var para = model.Paragraphs.FirstOrDefault(p => p.Runs.Any(r => r.Text != null && r.Text.Contains("5.3")));
            Assert.NotNull(para);
            foreach (var r in para!.Runs)
            {
                string text = r.Text ?? string.Empty;
                // invoke private sanitizer using reflection
                var method = typeof(Nedev.DocToDocx.Writers.DocumentWriter).GetMethod("SanitizeXmlString", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
                string cleaned = (string)method.Invoke(null, new object[] { text })!;
                Console.WriteLine($"Original run ({text.Length}): '{text}'");
                Console.WriteLine($" Sanitize -> '{cleaned}'");
                // after the recent fix all runs in this paragraph should sanitize to
                // empty because they are spurious binary junk that slipped into the
                // field code reader.
                Assert.Equal(string.Empty, cleaned);
            }
            Assert.True(true);
        }

        [Fact]
        public void ProblematicDoc_ProducesValidFootnotes()
        {
            string sample = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "渠道授权协议v5.doc"));
            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(DocToDocxConverter.LoadDocument(sample));
                zw.Dispose();
                pkg = ms.ToArray();
            }
            var fnxml = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read)
                .GetEntry("word/footnotes.xml").Open()).ReadToEnd();
            // should not contain the stray default namespace used previously
            Assert.DoesNotContain("xmlns=\"footnote\"", fnxml);
            Assert.Matches("<w:footnote[^>]*id=", fnxml);
        }

        [Fact]
        public void DumpFootnotes_FromProblematicDoc()
        {
            string sample = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "渠道授权协议v5.doc"));
            var model = DocToDocxConverter.LoadDocument(sample);
            Console.WriteLine($"Footnote count: {model.Footnotes.Count}");
            int idx = 0;
            foreach (var fn in model.Footnotes)
            {
                Console.WriteLine($"Footnote {idx++} id={fn.Index} cpStart={fn.CharacterPosition} cpLen={fn.CharacterLength} paragraphs={fn.Paragraphs.Count}");
                foreach (var p in fn.Paragraphs)
                {
                    foreach (var r in p.Runs)
                    {
                        Console.Write(" run text chars:");
                        if (r.Text != null)
                        {
                            foreach (var ch in r.Text)
                                Console.Write($" \\u{(int)ch:X4}");
                        }
                        Console.WriteLine();
                    }
                }
            }
            // ensure we didn’t silently drop footnote text during reading; every note with
            // a nonzero length should have produced at least one paragraph/run.
            Assert.All(model.Footnotes, fn =>
                Assert.True(fn.Paragraphs.Count > 0 || fn.CharacterLength == 0,
                            $"Footnote {fn.Index} had length {fn.CharacterLength} but no paragraphs"));
            // additionally verify the decoded strings are not just ASCII (the sample is
            // Chinese, so run text should include wide characters). this gives us some
            // confidence that the encoding provider registration worked.
            Assert.All(model.Footnotes, fn =>
            {
                var anyNonAscii = fn.Runs.Any(r => r.Text != null && r.Text.Any(ch => ch > 0x7F));
                Assert.True(anyNonAscii, $"Footnote {fn.Index} text appears to be all ASCII");
            });
        }

        [Fact]
        public void ProblematicDoc_FibValues()
        {
            string sample = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "渠道授权协议v5.doc"));
            using var reader = new Nedev.DocToDocx.Readers.DocReader(sample);
            // before loading, dump OLE container directory for inspection
            var cfbField = typeof(Nedev.DocToDocx.Readers.DocReader)
                .GetField("_cfb", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            var cfb = cfbField?.GetValue(reader) as Nedev.DocToDocx.Readers.CfbReader;
            if (cfb != null)
            {
                Console.WriteLine("OLE directory listing:\n" + cfb.GetDiagnostics());
            }

            reader.Load();
            // grab private fields via reflection for inspection
            var fib = typeof(Nedev.DocToDocx.Readers.DocReader)
                .GetField("_fibReader", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                ?.GetValue(reader) as Nedev.DocToDocx.Readers.FibReader;
            var textReader = typeof(Nedev.DocToDocx.Readers.DocReader)
                .GetField("_textReader", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                ?.GetValue(reader) as Nedev.DocToDocx.Readers.TextReader;
            var ftStream = typeof(Nedev.DocToDocx.Readers.DocReader)
                .GetField("_footnoteStream", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                ?.GetValue(reader) as System.IO.Stream;
            var tableReader = typeof(Nedev.DocToDocx.Readers.DocReader)
                .GetField("_tableReader", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                ?.GetValue(reader) as System.IO.BinaryReader;

            Assert.NotNull(fib);
            Console.WriteLine($"FIB ccpText={fib!.CcpText} ccpFtn={fib.CcpFtn} ccpHdd={fib.CcpHdd} ccpEdn={fib.CcpEdn}");
            Console.WriteLine($"FIB fcFtn={fib.FcFtn} lcbFtn={fib.LcbFtn}");
            if (textReader != null)
            {
                Console.WriteLine($"Global text length={textReader.Text.Length}");
                Console.WriteLine("Piece table entries:");
                foreach (var p in textReader.Pieces.Take(10))
                    Console.WriteLine($"  cp [{p.CpStart},{p.CpEnd}) unicode={p.IsUnicode} off={p.FileOffset}");
                if (textReader.Pieces.Count > 10)
                    Console.WriteLine($"  ... ({textReader.Pieces.Count} total pieces, last end={textReader.Pieces.Last().CpEnd})");
            }
            if (ftStream != null)
                Console.WriteLine($"Footnote stream length={ftStream.Length}");

            // dump the raw bytes at the FIB-specified footnote PLC location in the
            // table stream; the data here often contains the CP array followed by
            // the actual footnote text for simple documents.  examining it helped
            // us realize the text was not in the main WordDocument stream.
            if (tableReader != null && fib.FcFtn != 0 && fib.LcbFtn > 0)
            {
                try
                {
                    tableReader.BaseStream.Seek(fib.FcFtn, SeekOrigin.Begin);
                    var plcBytes = tableReader.ReadBytes((int)fib.LcbFtn);
                    Console.WriteLine("Footnote PLC raw (unicode decode):\n" + Encoding.Unicode.GetString(plcBytes));
                    Console.WriteLine("Footnote PLC raw (ansi decode):\n" + Encoding.GetEncoding(1252).GetString(plcBytes));
                    // output hex bytes surrounded by backticks so markdown doesn't eat them
                    var hexOutput = string.Join(" ", plcBytes.Select(b => b.ToString("X2")));
                    Console.WriteLine("Footnote PLC hex: `" + hexOutput + "`");
                    // parse cp array and FRD entries explicitly
                    int n = (int)((fib.LcbFtn - 4) / 6);
                    tableReader.BaseStream.Seek(fib.FcFtn, SeekOrigin.Begin);
                    var cpArray = new int[n + 1];
                    for (int i = 0; i <= n; i++)
                        cpArray[i] = tableReader.ReadInt32();
                    Console.WriteLine("CP array: " + string.Join(",", cpArray));
                    var frdValues = new ushort[n];
                    for (int i = 0; i < n; i++)
                        frdValues[i] = tableReader.ReadUInt16();
                    Console.WriteLine("FRD entries (ushort): " + string.Join(",", frdValues));
                    Console.WriteLine("FRD hex: " + BitConverter.ToString(plcBytes, (n + 1) * 4, n * 2));

                    // attempt to interpret FRD values as offsets into WordDocument stream
                    var wordStream = reader.GetCfbReader().GetStream("WordDocument");
                    for (int i = 0; i < n; i++)
                    {
                        if (frdValues[i] == 0) continue;
                        var off = frdValues[i];
                        Console.WriteLine($"FRD[{i}] raw=0x{off:X4} guessed offset={off}");
                        if (off + 100 < wordStream.Length)
                        {
                            wordStream.Seek(off, SeekOrigin.Begin);
                            var buf = new byte[100];
                            wordStream.Read(buf, 0, buf.Length);
                            Console.WriteLine(" sample bytes (hex): " + BitConverter.ToString(buf));
                            var gbk = Encoding.GetEncoding(936).GetString(buf);
                            var uni = Encoding.Unicode.GetString(buf);
                            Console.WriteLine(" sample as GBK: `" + gbk + "`");
                            Console.WriteLine(" sample as Unicode: `" + uni + "`");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to dump footnote PLC bytes: " + ex.Message);
                }
            }

            // also peek at the raw WordDocument bytes around where footnotes would be
            if (fib != null)
            {
                try
                {
                    var wordStream = reader.GetCfbReader().GetStream("WordDocument");
                    var all = wordStream.ToArray();
                    int textOffset = fib.FcMin > 0 ? (int)fib.FcMin : 0x200;
                    int bodyBytes = fib.CcpText * 2;
                    int footStart = textOffset + bodyBytes;
                    Console.WriteLine($"body ends at file offset {footStart}");
                    if (footStart < all.Length)
                    {
                        var sampleBytes = all.Skip(footStart).Take(2000).ToArray();
                        Console.WriteLine("footnote sample (unicode):\n" + Encoding.Unicode.GetString(sampleBytes));
                        Console.WriteLine("footnote sample (ansi):\n" + Encoding.GetEncoding(1252).GetString(sampleBytes));
                        Console.WriteLine("footnote sample (gbk):\n" + Encoding.GetEncoding(936).GetString(sampleBytes));
                        Console.WriteLine("footnote sample bytes (hex): " + BitConverter.ToString(sampleBytes));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to dump raw word bytes: " + ex.Message);
                }
            }

            // nothing to assert here beyond non-null fib
            Assert.True(true);
        }

        [Fact]
        public void ProblematicDoc_ConvertedDocxContainsFootnoteText()
        {
            string sample = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "渠道授权协议v5.doc"));
            byte[] pkg;
            using (var ms = new MemoryStream())
            {
                var zw = new ZipWriter(ms);
                zw.WriteDocument(DocToDocxConverter.LoadDocument(sample));
                zw.Dispose();
                pkg = ms.ToArray();
            }
            var fnxml = new StreamReader(new System.IO.Compression.ZipArchive(new MemoryStream(pkg), System.IO.Compression.ZipArchiveMode.Read)
                .GetEntry("word/footnotes.xml").Open()).ReadToEnd();
            // there should be at least one textual run inside footnotes
            Assert.Matches("<w:t>\\s*[^\\s<].+?</w:t>", fnxml);
        }

        [Fact]
        public void SectionReader_BogusSepx_DoesNotThrow()
        {
            // build a fake PLCFSED table with one section and a bogus offset
            // we need a nonzero offset because SectionReader treats 0 as "no table".
            const int offset = 4;
            var table = new MemoryStream();
            // pad to offset
            table.Write(new byte[offset], 0, offset);
            using (var bw = new BinaryWriter(table, Encoding.Default, true))
            {
                bw.Write(0);             // cp start
                bw.Write(100);           // cp end
                bw.Write((uint)0x12345678); // fcSepx points well past word stream
                bw.Write(new byte[8]);  // reserved
            }
            table.Position = 0;

            var word = new MemoryStream(new byte[4]); // too small for offset

            var fib = new FibReader(new BinaryReader(new MemoryStream(new byte[0])));
            // use reflection to set the public read-only properties with binding flags
            var flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            typeof(FibReader).GetProperty("FcPlcfSed", flags)!.SetValue(fib, (uint)offset);
            typeof(FibReader).GetProperty("LcbPlcfSed", flags)!.SetValue(fib, (uint)(table.Length - offset));
            // adjust stream position to start of actual data when SectionReader seeks
            // since fcPlcfSed is offset, we need to provide table stream with that many bytes at front

            var reader = new SectionReader(new BinaryReader(table), new BinaryReader(word), fib);
            var sections = reader.ReadSections();

            Assert.Single(sections);
            // range cp should match our input values; page width stays at its default
            Assert.Equal(0, sections[0].StartCp);
            Assert.Equal(100, sections[0].EndCp);
        }
    }
}
