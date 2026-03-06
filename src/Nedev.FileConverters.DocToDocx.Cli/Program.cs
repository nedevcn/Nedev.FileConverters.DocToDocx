using System;
using System.IO;
using System.Threading.Tasks;
using Nedev.FileConverters.DocToDocx;

namespace Nedev.FileConverters.DocToDocx.Cli;

public class Program
{
    public static async Task Main(string[] args)
    {
        Console.WriteLine("Nedev.DocToDocx CLI converter");
        Console.WriteLine("=============================");

        if (args.Length == 0 || args[0] == "-h" || args[0] == "--help")
        {
            PrintUsage();
            return;
        }

        if (args.Length == 1 && (args[0] == "-v" || args[0] == "--version"))
        {
            var ver = typeof(Program).Assembly.GetName().Version;
            Console.WriteLine($"Version {ver}");
            return;
        }

        if (args.Length < 2)
        {
            Console.WriteLine("Error: Missing required arguments.");
            PrintUsage();
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];
        string? password = null;
        bool recursive = false;
        bool disableHyperlinks = false;

        // parse remaining options
        for (int i = 2; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-p":
                case "--password":
                    if (i + 1 < args.Length)
                    {
                        password = args[++i];
                    }
                    break;
                case "-r":
                case "--recursive":
                    recursive = true;
                    break;
                case "--no-hyperlinks":
                    disableHyperlinks = true;
                    break;
            }
        }

        try
        {
            if (Directory.Exists(inputPath))
            {
                if (!Directory.Exists(outputPath))
                    Directory.CreateDirectory(outputPath);

                var search = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                    var files = Directory.GetFiles(inputPath, "*.*", search)
                                         .Where(f => f.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)
                                                  || f.EndsWith(".docx", StringComparison.OrdinalIgnoreCase));
                    foreach (var docFile in files)
                    {
                        var rel = Path.GetRelativePath(inputPath, docFile);
                        var outFile = Path.Combine(outputPath, Path.ChangeExtension(rel, ".docx"));
                        var outDir = Path.GetDirectoryName(outFile);
                        if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                            Directory.CreateDirectory(outDir);

                        if (docFile.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"Copying {docFile} -> {outFile}");
                            File.Copy(docFile, outFile, overwrite: true);
                        }
                        else
                        {
                            Console.WriteLine($"Converting {docFile} -> {outFile}");
                            DocToDocxConverter.Convert(docFile, outFile, password, enableHyperlinks: !disableHyperlinks);
                        }
                    }

                    Console.WriteLine("Directory conversion complete.");
                    return;
                }

                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file not found: {inputPath}");
                    return;
                }

                Console.WriteLine($"Input:  {inputPath}");
                Console.WriteLine($"Output: {outputPath}");
                Console.WriteLine("Converting...");

                var progress = new Progress<ConversionProgress>(p =>
                {
                    if (!string.IsNullOrEmpty(p.Message))
                    {
                        Console.WriteLine($"[{p.PercentComplete,3}%] {p.Stage}: {p.Message}");
                    }
                    else
                    {
                        Console.WriteLine($"[{p.PercentComplete,3}%] {p.Stage}");
                    }
                });

                await Task.Run(() => DocToDocxConverter.Convert(inputPath, outputPath, progress, password, enableHyperlinks: !disableHyperlinks));

            Console.WriteLine("Successfully converted the document.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            Environment.ExitCode = 1;
        }
    }

    static void PrintUsage()
    {
        Console.WriteLine("Usage: Nedev.DocToDocx.Cli <input.doc|inputDir> <output.docx|outputDir> [-p <password>] [-r]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  <input.doc>      The path to the input MS-DOC file.");
        Console.WriteLine("  <output.docx>    The path where the output DOCX file will be saved.");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -p, --password   The password to open an encrypted DOC file.");
        Console.WriteLine("  -r, --recursive  When input is a directory, process .doc files recursively.");
        Console.WriteLine("  -h, --help       Show this help message and exit.");
        Console.WriteLine("      --no-hyperlinks    Disable hyperlink elements and relationships in the generated DOCX.\n" +
                          "                       Text remains but links are treated as regular text, avoiding Word warnings");
    }
}
