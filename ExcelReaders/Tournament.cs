using System;
using Microsoft.Office.Interop.Excel;

namespace TournamentCalculator.ExcelReaders
{
    public class Tournament
    {
        public static dynamic IsWinnerDecided(Worksheet worksheet)
        {
            return !String.IsNullOrEmpty(Convert.ToString(worksheet.Range["DT41", Type.Missing].Value2));
        }

        public static bool IsGroupStageFinished(Worksheet worksheet)
        {
            return !String.IsNullOrEmpty(Convert.ToString(worksheet.Range["F45", Type.Missing].Value2));
        }
    }
}
