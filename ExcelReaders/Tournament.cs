using System;
using OfficeOpenXml;

namespace TournamentCalculator.ExcelReaders
{
    public class Tournament
    {
        public static dynamic IsWinnerDecided(ExcelWorksheet worksheet)
        {
            return !String.IsNullOrEmpty(Convert.ToString(worksheet.Cells["DT41"].Value));
        }

        public static bool IsGroupStageFinished(ExcelWorksheet worksheet)
        {
            var lastMatch = Convert.ToString(worksheet.Cells["F45"].Value);
            return !String.IsNullOrEmpty(lastMatch);
        }
    }
}