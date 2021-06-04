using System;
using OfficeOpenXml;

namespace TournamentCalculator.ExcelReaders
{
    public class Tournament
    {
        public static dynamic IsWinnerDecided(ExcelWorksheet worksheet)
        {
            return !String.IsNullOrEmpty(Convert.ToString(worksheet.Cells["FO41"].Value));
        }

        //public static bool IsBronzeWinnerDecided(ExcelWorksheet worksheet)
        //{
        //    var lastMatch = Convert.ToString(worksheet.Cells["BS35"].Value);
        //    return !String.IsNullOrEmpty(lastMatch);
        //}

        public static bool IsGroupStageFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["F45"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }

        public static bool IsEightFinalsFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["EX39"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }

        public static bool IsQuarterFinalsFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["FE37"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }

        public static bool IsSemiFinalsFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["FL33"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }
    }
}