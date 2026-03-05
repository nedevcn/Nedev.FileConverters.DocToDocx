using System;
using System.IO;
using System.Text.RegularExpressions;

var path = @"d:\Project\DocToDocx\src\Readers\DocReader.cs";
var text = File.ReadAllText(path);

var match = Regex.Match(text, @"private static void TryPopulateChartFromSourceBytes\s*\(\s*ChartModel\s*model\s*\)\s*\{");
if (match.Success)
{
    int startIdx = match.Index;
    int braceCount = 0;
    int endIdx = -1;
    bool foundFirstBrace = false;

    for (int i = startIdx; i < text.Length; i++)
    {
        if (text[i] == '{')
        {
            braceCount++;
            foundFirstBrace = true;
        }
        else if (text[i] == '}')
        {
            braceCount--;
            if (foundFirstBrace && braceCount == 0)
            {
                endIdx = i;
                break;
            }
        }
    }

    if (endIdx != -1)
    {
        int summaryIdx = text.LastIndexOf("/// <summary>", startIdx);
        if (summaryIdx != -1 && startIdx - summaryIdx < 500) {
            startIdx = summaryIdx;
        }

        var newText = text.Substring(0, startIdx) + text.Substring(endIdx + 1);
        File.WriteAllText(path, newText);
        Console.WriteLine("Successfully removed TryPopulateChartFromSourceBytes.");
    }
}
