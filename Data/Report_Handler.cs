using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Windows;
using Microsoft.Win32;
using System.IO;

namespace PMBusInterface.Data
{
    static class Report_Handler
    {
        private static readonly string _defaultFileName = @"\Report";
        private static readonly string _defaultFileExt = ".xlsx";

        public static void CreateReport(string filePath)
        {
            ExcelPackage.License.SetNonCommercialOrganization("BitBlaze");

            string fullPath = filePath + _defaultFileName + _defaultFileExt;

            FileInfo newFile = new FileInfo(fullPath);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(fullPath);
            }
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                // add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Inventory");

                package.Save();
                MessageBox.Show("File successfully saved. Path: " + fullPath);
            }
        }

        public static string GetDefaultFilePath()
        {
            string userName = Environment.UserName;
            return @"C:\Users\" + userName + @"\Desktop";
        }
    }
}
