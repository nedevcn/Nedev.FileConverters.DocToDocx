#nullable enable
using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Nedev.DocToDocx;
using Nedev.DocToDocx.Cli;
using Nedev.DocToDocx.Models;

namespace Nedev.DocToDocx.Tests
{
    public class ReaderIntegrationTests
    {
        [Fact]
        public void CreateAndLoadDocument_HasContent()
        {
            // Build a simple document in memory and save it as DOCX, then load via API.
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            var original = new DocumentModel();
            original.Paragraphs.Add(new ParagraphModel { Runs = { new RunModel { Text = "Hello" } } });
            DocToDocxConverter.SaveDocument(original, path);

            var doc = DocToDocxConverter.LoadDocument(path);
            Assert.NotNull(doc);
            Assert.Equal(1, doc.Paragraphs.Count);
            File.Delete(path);
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
            string inPath = Path.Combine(AppContext.BaseDirectory, "test.doc");
            string outPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");

            var reports = new List<ConversionProgress>();
            var progress = new Progress<ConversionProgress>(p => reports.Add(p));

            DocToDocxConverter.Convert(inPath, outPath, progress);

            Assert.Contains(reports, r => r.Stage == ConversionStage.Reading);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Writing);
            Assert.Contains(reports, r => r.Stage == ConversionStage.Complete);

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

            using var zip = new System.IO.Compression.ZipArchive(File.OpenRead(tempOut), System.IO.Compression.ZipArchiveMode.Read);
            var entry = zip.GetEntry("word/document.xml");
            Assert.NotNull(entry);
            using var reader = new StreamReader(entry.Open());
            var xml = reader.ReadToEnd();
            Assert.Contains("Hello", xml);

            File.Delete(tempOut);
        }
    }
}