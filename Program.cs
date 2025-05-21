using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using HtmlAgilityPack;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length != 2)
        {
            Console.WriteLine("Usage: ConvertHtmlToExcel <input.xls> <output.xlsx>");
            return;
        }

        var inputPath = args[0];
        var outputPath = args[1];

        if (!File.Exists(inputPath))
        {
            Console.WriteLine("Input file not found.");
            return;
        }

        var html = File.ReadAllText(inputPath);
        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        var table = doc.DocumentNode.SelectSingleNode("//table");
        if (table == null)
        {
            Console.WriteLine("No <table> found in HTML.");
            return;
        }

        using var workbook = new XLWorkbook();
        var sheetName = Path.GetFileNameWithoutExtension(outputPath); 
        var worksheet = workbook.Worksheets.Add(sheetName); 
    
        int row = 1;

        foreach (var tr in table.SelectNodes(".//tr"))
        {
            int col = 1;
            foreach (var cell in tr.Elements("th").Concat(tr.Elements("td")))
            {
                worksheet.Cell(row, col++).Value = cell.InnerText.Trim();
            }
            row++;
        }

        workbook.SaveAs(outputPath);
        Console.WriteLine($"File saved to: {outputPath}");
    }
}
