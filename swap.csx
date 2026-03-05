using System;
using System.IO;

var path = @"d:\Project\DocToDocx\src\Readers\DocReader.cs";
var text = File.ReadAllText(path);
text = text.Replace("TryPopulateChartFromSourceBytes(model);", "BiffChartScanner.TryPopulateChart(model);");
File.WriteAllText(path, text);
Console.WriteLine("Replaced call");
