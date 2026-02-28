# ConvertHtmlToExcel

## Purpose

The main purpose of this utility is to convert incorrectly generated `.xls` files—which actually just contain an HTML `<table>` inside—into proper, standard `.xlsx` files.

Many older or poorly implemented systems export data as "Excel" by simply generating an HTML table and saving it with an `.xls` extension. When opening these files in modern spreadsheet applications, you often get a warning prompt about the file format not matching the extension. This tool parses the underlying HTML, extracts the table data, and generates a clean, native Excel workbook (`.xlsx`).

## How It Works

This is a C# Console Application that uses:
- **HtmlAgilityPack** to load and parse the HTML structure of the input `.xls` file.
- **ClosedXML** to generate and save the new native `.xlsx` workbook.

It searches for the `<table>` element, iterates through its `<tr>` (rows), `<th>` (headers), and `<td>` (cells), and writes their inner text sequentially to a new excel worksheet.

## Usage

Run the compiled executable from the command line, providing the input "fake" `.xls` file and the desired output `.xlsx` file path:

```sh
ConvertHtmlToExcel <input.xls> <output.xlsx>
```

### Example

```sh
ConvertHtmlToExcel Contacts.xls Contacts.xlsx
```

## Requirements

- .NET Runtime
- Dependencies are managed via NuGet (`ClosedXML`, `HtmlAgilityPack`).
