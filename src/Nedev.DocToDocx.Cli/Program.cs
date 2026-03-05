using System;
using System.IO;
using System.Threading.Tasks;
using Nedev.DocToDocx;

namespace Nedev.DocToDocx.Cli;

class Program
{
    static async Task Main(string[] args)
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
        string? password = null;

        // Parse optional password argument
        if (args.Length >= 4 && (args[2] == "-p" || args[2] == "--password"))
        {
            password = args[3];
        }

        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found: {inputPath}");
            return;
        }

        try
        {
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

            // Using the progress-enabled synchronous Convert method, but we could also use ConvertAsync if needed.
            await Task.Run(() => DocToDocxConverter.Convert(inputPath, outputPath, progress, password));

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
        Console.WriteLine("Usage: Nedev.DocToDocx.Cli <input.doc> <output.docx> [-p <password>]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  <input.doc>      The path to the input MS-DOC file.");
        Console.WriteLine("  <output.docx>    The path where the output DOCX file will be saved.");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -p, --password   The password to open an encrypted DOC file.");
        Console.WriteLine("  -h, --help       Show this help message and exit.");
    }
}
