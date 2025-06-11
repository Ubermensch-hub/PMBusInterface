using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace PMBusInterface.Data
{
    static class Report_Handler
    {
        private static readonly string _defaultFileName = @"\Report";
        private static readonly string _defaultFileExt = ".xlsx";

        public static class ReportParser
        {
            public static Dictionary<string, string> ParseReport(string report)
            {
                var result = new Dictionary<string, string>();
                var lines = report.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var line in lines)
                {
                    // Обрабатываем строки вида "Параметр: значение"
                    var match = Regex.Match(line, @"^(.*?):(.*)$");
                    if (match.Success)
                    {
                        var key = match.Groups[1].Value.Trim();
                        var value = match.Groups[2].Value.Trim();
                        result[key] = value;
                    }
                }

                return result;
            }
        }

        public static void CreateReport(string filePath, string psuData, string fruData)
        {
            ExcelPackage.License.SetNonCommercialOrganization("BitBlaze");

            string fullPath = filePath + _defaultFileName + _defaultFileExt;
            var fileInfo = new FileInfo(fullPath);
            if (fileInfo.Exists) fileInfo.Delete();

            using (var package = new ExcelPackage(fileInfo))
            {
                // Лист для PSU отчета
                var psuWorksheet = package.Workbook.Worksheets.Add("PSU Report");
                ExportStructuredData(psuWorksheet, ReportParser.ParseReport(psuData), "PSU Report");

                // Лист для FRU отчета (аналогично)
                var fruWorksheet = package.Workbook.Worksheets.Add("FRU Report");
                ExportStructuredData(fruWorksheet, ReportParser.ParseReport(fruData), "FRU Report");

                package.Save();
                MessageBox.Show($"Отчёт сохранён: {fullPath}");
            }
        }

        private static void ExportStructuredData(ExcelWorksheet worksheet,
                                      Dictionary<string, string> data,
                                      string reportTitle)
        {
            // Стиль границ
            var borderStyle = ExcelBorderStyle.Thin;
            var borderColor = System.Drawing.Color.Black;

            // Заголовок отчета
            worksheet.Cells["A1"].Value = reportTitle;
            worksheet.Cells["A1:B1"].Merge = true;
            worksheet.Cells["A1"].Style.Font.Bold = true;
            worksheet.Cells["A1"].Style.Font.Size = 14;

            // Заголовки колонок
            worksheet.Cells["A2"].Value = "Parameter";
            worksheet.Cells["B2"].Value = "Value";
            worksheet.Cells["A2:B2"].Style.Font.Bold = true;

            ApplyAllBorders(worksheet.Cells["A2:B2"], borderStyle, borderColor);

            // Заполняем данные
            int row = 3;
            foreach (var item in data)
            {
                worksheet.Cells[$"A{row}"].Value = item.Key;
                worksheet.Cells[$"B{row}"].Value = item.Value;
                ApplyAllBorders(worksheet.Cells[$"A{row}:B{row}"], borderStyle, borderColor);

                row++;
            }

            // Форматирование
            worksheet.Cells["A:B"].AutoFitColumns();
            worksheet.Cells["A2:B2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Cells["A2:B2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
        }

        private static void ApplyAllBorders(ExcelRange cells, ExcelBorderStyle style, System.Drawing.Color color)
        {
            cells.Style.Border.Top.Style = style;
            cells.Style.Border.Top.Color.SetColor(color);
            cells.Style.Border.Bottom.Style = style;
            cells.Style.Border.Bottom.Color.SetColor(color);
            cells.Style.Border.Left.Style = style;
            cells.Style.Border.Left.Color.SetColor(color);
            cells.Style.Border.Right.Style = style;
            cells.Style.Border.Right.Color.SetColor(color);
        }

        public static string GetDefaultFilePath()
        {
            string userName = Environment.UserName;
            return @"C:\Users\" + userName + @"\Desktop";
        }
    }
}
