using System.IO;
using System.Text.RegularExpressions;

var path = @"d:\Project\DocToDocx\src\Readers\DocReader.cs";
var text = File.ReadAllText(path);

// Find the start of the TableReader class
var match = Regex.Match(text, @"summary>\s*public class TableReader\s*\{");
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
        // Find if there are preceding summary comments to remove too
        var summaryMatch = Regex.Match(text.Substring(0, startIdx), @"///\s*<summary>.*?///\s*</summary>\s*$" , RegexOptions.Singleline);
        if (summaryMatch.Success) {
            startIdx = summaryMatch.Index;
        }

        var newText = text.Substring(0, startIdx) + text.Substring(endIdx + 1);
        File.WriteAllText(path, newText);
    }
}
