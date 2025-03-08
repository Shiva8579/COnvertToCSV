using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using System.Text;

class Program
{
    static void Main()
    {
        string excelPath = @"D:\MigrationCSV\XpressData_28Feb2025\Transactions status.xlsx";
        string folderPath = @"D:\MigrationCSV\XpressData_28Feb2025\20250303\";

        using (var workbook = new XLWorkbook(excelPath))
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                // Define the CSV path based on sheet name
                string csvPath = Path.Combine(folderPath, $"{worksheet.Name}.csv");

                var rows = worksheet.RangeUsed().RowsUsed(); 
                int totalColumns = worksheet.RangeUsed().ColumnCount();

                using (StreamWriter writer = new StreamWriter(csvPath, false, Encoding.UTF8))
                {
                    foreach (var row in rows)
                    {
                        string[] rowValues = new string[totalColumns];

                        for (int col = 1; col <= totalColumns; col++)
                        {
                            var cellValue = row.Cell(col).Value.ToString().Trim();

                            // Preserve empty values and wrap values with commas in quotes
                            if (string.IsNullOrEmpty(cellValue))
                            {
                                rowValues[col - 1] = "";
                            }
                            else if (cellValue.Contains(",") || cellValue.Contains("\""))
                            {
                                // Escape double quotes by replacing `"` with `""`
                                rowValues[col - 1] = $"\"{cellValue.Replace("\"", "\"\"")}\"";
                            }
                            else
                            {
                                rowValues[col - 1] = cellValue;
                            }
                        }

                        // Write the row to the CSV file
                        writer.WriteLine(string.Join(",", rowValues));
                    }
                }

                Console.WriteLine($"Sheet '{worksheet.Name}' converted to CSV.");
            }
        }

        Console.WriteLine("All sheets converted to CSV.");
    }
}
