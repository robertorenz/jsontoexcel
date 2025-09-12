using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            ShowHelp();
            return;
        }

        // Parse arguments
        string inputFile = "";
        string outputFile = "";
        bool enableFormatting = true; // Default to formatting enabled

        // Process arguments
        var argsList = new List<string>(args);
        
        // Check for --no-format flag
        if (argsList.Contains("--no-format"))
        {
            enableFormatting = false;
            argsList.Remove("--no-format");
        }
        
        // Check for help flags
        if (argsList.Contains("--help") || argsList.Contains("-h"))
        {
            ShowHelp();
            return;
        }

        if (argsList.Count < 2)
        {
            ShowHelp();
            return;
        }

        inputFile = argsList[0];
        outputFile = argsList[1];

        try
        {
            ConvertJsonToExcel(inputFile, outputFile, enableFormatting);
            string formatStatus = enableFormatting ? "with formatting" : "without formatting";
            Console.WriteLine($"Successfully converted {inputFile} to {outputFile} {formatStatus}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static void ShowHelp()
    {
        Console.WriteLine("JsonToExcel - Convert JSON files to Excel format");
        Console.WriteLine();
        Console.WriteLine("Usage: JsonToExcel <input.json> <output.xlsx> [options]");
        Console.WriteLine();
        Console.WriteLine("Arguments:");
        Console.WriteLine("  input.json    Path to the input JSON file");
        Console.WriteLine("  output.xlsx   Path to the output Excel file");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --no-format   Disable formatting (creates plain Excel output)");
        Console.WriteLine("  --help, -h    Show this help message");
        Console.WriteLine();
        Console.WriteLine("Examples:");
        Console.WriteLine("  JsonToExcel data.json output.xlsx");
        Console.WriteLine("  JsonToExcel data.json output.xlsx --no-format");
        Console.WriteLine();
        Console.WriteLine("Features with formatting enabled:");
        Console.WriteLine("  • Professional blue headers with white text");
        Console.WriteLine("  • Alternating row colors for readability");
        Console.WriteLine("  • Data type-specific formatting (numbers, dates, emails)");
        Console.WriteLine("  • Auto-sized columns with borders");
        Console.WriteLine("  • Color-coded boolean values (green/red)");
    }

    static void ConvertJsonToExcel(string jsonFile, string excelFile, bool enableFormatting = true)
    {
        if (!File.Exists(jsonFile))
        {
            throw new FileNotFoundException($"JSON file not found: {jsonFile}");
        }

        string jsonContent = File.ReadAllText(jsonFile);
        JToken jsonToken = JToken.Parse(jsonContent);

        // Set EPPlus license for noncommercial use
        ExcelPackage.License.SetNonCommercialPersonal("JsonToExcel App User");

        using (var package = new ExcelPackage())
        {
            var worksheet = package.Workbook.Worksheets.Add("Data");

            if (jsonToken is JArray jsonArray)
            {
                WriteArrayToWorksheet(jsonArray, worksheet, enableFormatting);
            }
            else if (jsonToken is JObject jsonObject)
            {
                WriteObjectToWorksheet(jsonObject, worksheet, enableFormatting);
            }
            else
            {
                worksheet.Cells[1, 1].Value = jsonToken.ToString();
            }

            FileInfo excelFileInfo = new FileInfo(excelFile);
            package.SaveAs(excelFileInfo);
        }
    }

    static void WriteArrayToWorksheet(JArray jsonArray, ExcelWorksheet worksheet, bool enableFormatting = true)
    {
        if (jsonArray.Count == 0) return;

        if (jsonArray[0] is JObject firstObject)
        {
            var properties = firstObject.Properties().ToArray();
            
            // Create header row
            for (int col = 0; col < properties.Length; col++)
            {
                var headerCell = worksheet.Cells[1, col + 1];
                headerCell.Value = properties[col].Name;
                
                if (enableFormatting)
                {
                    // Header formatting
                    headerCell.Style.Font.Bold = true;
                    headerCell.Style.Font.Color.SetColor(Color.White);
                    headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
                    headerCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    headerCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    headerCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    headerCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headerCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
                else
                {
                    // Basic header formatting (just bold)
                    headerCell.Style.Font.Bold = true;
                }
            }

            // Add data rows
            for (int row = 0; row < jsonArray.Count; row++)
            {
                if (jsonArray[row] is JObject obj)
                {
                    for (int col = 0; col < properties.Length; col++)
                    {
                        var dataCell = worksheet.Cells[row + 2, col + 1];
                        var value = obj[properties[col].Name];
                        dataCell.Value = ConvertValue(value);
                        
                        if (enableFormatting)
                        {
                            // Data cell formatting with alternating colors
                            bool isEvenRow = row % 2 == 0;
                            Color rowColor = isEvenRow ? Color.FromArgb(242, 242, 242) : Color.White;
                            
                            dataCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            dataCell.Style.Fill.BackgroundColor.SetColor(rowColor);
                            dataCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            dataCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            dataCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            dataCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            dataCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.LightGray);
                            
                            // Format numbers and dates
                            FormatDataCell(dataCell, value);
                        }
                    }
                }
            }
            
            // Auto-fit columns
            for (int col = 1; col <= properties.Length; col++)
            {
                worksheet.Column(col).AutoFit();
                if (enableFormatting && worksheet.Column(col).Width > 50) // Max width limit
                    worksheet.Column(col).Width = 50;
            }
        }
        else
        {
            // Handle simple array
            var headerCell = worksheet.Cells[1, 1];
            headerCell.Value = "Value";
            
            if (enableFormatting)
            {
                // Header formatting
                headerCell.Style.Font.Bold = true;
                headerCell.Style.Font.Color.SetColor(Color.White);
                headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
                headerCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            else
            {
                // Basic header formatting (just bold)
                headerCell.Style.Font.Bold = true;
            }
            
            for (int i = 0; i < jsonArray.Count; i++)
            {
                var dataCell = worksheet.Cells[i + 2, 1];
                dataCell.Value = ConvertValue(jsonArray[i]);
                
                if (enableFormatting)
                {
                    // Alternating row colors
                    bool isEvenRow = i % 2 == 0;
                    Color rowColor = isEvenRow ? Color.FromArgb(242, 242, 242) : Color.White;
                    dataCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    dataCell.Style.Fill.BackgroundColor.SetColor(rowColor);
                    dataCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.LightGray);
                    
                    FormatDataCell(dataCell, jsonArray[i]);
                }
            }
            
            worksheet.Column(1).AutoFit();
        }
    }

    static void WriteObjectToWorksheet(JObject jsonObject, ExcelWorksheet worksheet, bool enableFormatting = true)
    {
        // Setup headers
        var propertyHeader = worksheet.Cells[1, 1];
        var valueHeader = worksheet.Cells[1, 2];
        
        propertyHeader.Value = "Property";
        valueHeader.Value = "Value";
        
        if (enableFormatting)
        {
            // Header formatting
            foreach (var headerCell in new[] { propertyHeader, valueHeader })
            {
                headerCell.Style.Font.Bold = true;
                headerCell.Style.Font.Color.SetColor(Color.White);
                headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerCell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(68, 114, 196));
                headerCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }
        }
        else
        {
            // Basic header formatting (just bold)
            foreach (var headerCell in new[] { propertyHeader, valueHeader })
            {
                headerCell.Style.Font.Bold = true;
            }
        }

        var properties = jsonObject.Properties().ToArray();
        for (int i = 0; i < properties.Length; i++)
        {
            var propertyCell = worksheet.Cells[i + 2, 1];
            var valueCell = worksheet.Cells[i + 2, 2];
            
            propertyCell.Value = properties[i].Name;
            valueCell.Value = ConvertValue(properties[i].Value);
            
            if (enableFormatting)
            {
                // Alternating row colors
                bool isEvenRow = i % 2 == 0;
                Color rowColor = isEvenRow ? Color.FromArgb(242, 242, 242) : Color.White;
                
                foreach (var dataCell in new[] { propertyCell, valueCell })
                {
                    dataCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    dataCell.Style.Fill.BackgroundColor.SetColor(rowColor);
                    dataCell.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.LightGray);
                }
                
                // Format the value cell based on data type
                FormatDataCell(valueCell, properties[i].Value);
            }
        }
        
        // Auto-fit columns
        worksheet.Column(1).AutoFit();
        worksheet.Column(2).AutoFit();
        if (enableFormatting && worksheet.Column(2).Width > 50)
            worksheet.Column(2).Width = 50;
    }

    static void FormatDataCell(ExcelRange cell, JToken token)
    {
        if (token == null || token.Type == JTokenType.Null)
            return;

        switch (token.Type)
        {
            case JTokenType.Integer:
                cell.Style.Numberformat.Format = "#,##0";
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                break;
            case JTokenType.Float:
                cell.Style.Numberformat.Format = "#,##0.00";
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                break;
            case JTokenType.Date:
                cell.Style.Numberformat.Format = "mm/dd/yyyy";
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                break;
            case JTokenType.Boolean:
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                // Color boolean values
                if (token.Value<bool>())
                {
                    cell.Style.Font.Color.SetColor(Color.Green);
                    cell.Style.Font.Bold = true;
                }
                else
                {
                    cell.Style.Font.Color.SetColor(Color.Red);
                }
                break;
            case JTokenType.String:
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                
                // Special formatting for email addresses
                string? stringValue = token.Value<string>();
                if (!string.IsNullOrEmpty(stringValue) && stringValue.Contains("@") && stringValue.Contains("."))
                {
                    cell.Style.Font.Color.SetColor(Color.Blue);
                    cell.Style.Font.UnderLine = true;
                }
                break;
            case JTokenType.Array:
            case JTokenType.Object:
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                cell.Style.Font.Italic = true;
                cell.Style.Font.Color.SetColor(Color.Gray);
                break;
        }
    }

    static object ConvertValue(JToken token)
    {
        if (token == null || token.Type == JTokenType.Null)
            return null;

        switch (token.Type)
        {
            case JTokenType.Integer:
                return token.Value<long>();
            case JTokenType.Float:
                return token.Value<double>();
            case JTokenType.Boolean:
                return token.Value<bool>();
            case JTokenType.Date:
                return token.Value<DateTime>();
            case JTokenType.String:
                return token.Value<string>();
            case JTokenType.Array:
            case JTokenType.Object:
                return token.ToString(Formatting.None);
            default:
                return token.ToString();
        }
    }
}
