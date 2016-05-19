using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TournamentCalculator
{
    public class ExcelService
    {
        public static Worksheet GetWorksheet(Application excel, string file)
        {
            var workbook = excel.Workbooks.Open(
                file, Missing.Value, false, Missing.Value, "", "", true,
                XlPlatform.xlWindows,
                Missing.Value, false, false, Missing.Value, false, Missing.Value, Missing.Value);

            // get the collection of sheets in the workbook
            var sheets = workbook.Worksheets;

            // get the correct worksheet from the collection of worksheets
            return (Worksheet)sheets.Item[4];   
        }

        public static void KillAllExcelProcesses()
        {
            var p = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            for (var i = 0; i < p.GetLength(0); i++)
                p[i].Kill();
        }
        public static string GetResultWorksheet(string fasitFile)
        {
            string[] files = Directory.GetFiles(fasitFile);
            if (files.Length == 0)
                throw new Exception("Fasit ikke funnet: " + fasitFile);

            return files[0];
        }

        public static void Cleanup(Application excel)
        {
            excel.Quit();

            Marshal.ReleaseComObject(excel);
            GC.Collect();

            KillAllExcelProcesses();
        }
    }
}