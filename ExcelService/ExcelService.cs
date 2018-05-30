using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace TournamentCalculator.ExcelService
{
    public class ExcelService
    {
        public static ExcelWorksheet GetWorksheet(string file)
        {
            var package = new ExcelPackage(new FileInfo(file));
            return package.Workbook.Worksheets.Last();
        }

        public static string GetResultWorksheet(string fasitFile)
        {
            var files = Directory.GetFiles(fasitFile);

            if (files.Length == 0)
                throw new Exception("Fasit ikke funnet: " + fasitFile);

            return files[0];
        }
    }
}