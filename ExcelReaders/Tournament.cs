using System;
using OfficeOpenXml;

namespace TournamentCalculator.ExcelReaders
{
    public class Tournament
    {
        public static dynamic IsWinnerDecided(ExcelWorksheet worksheet)
        {
            return !String.IsNullOrEmpty(Convert.ToString(worksheet.Cells["BO41"].Value));
        }

        public static bool IsBronzeWinnerDecided(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["BS35"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }

        public static bool IsGroupStageFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["F54"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }

        public static bool IsEightFinalsFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["BA39"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }

        public static bool IsQuarterFinalsFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["BG37"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }

        public static bool IsSemiFinalsFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["BM33"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }
    }
}