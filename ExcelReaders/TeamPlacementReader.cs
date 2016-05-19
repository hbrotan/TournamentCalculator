using System;
using System.Collections.Specialized;
using Microsoft.Office.Interop.Excel;

namespace TournamentCalculator.ExcelReaders
{
    public class TeamPlacementReader
    {
        public static string GetWinner(Worksheet worksheet)
        {
            var val = Convert.ToString(worksheet.Range["DT41", Type.Missing].Value2);          
            return !String.IsNullOrEmpty(val) ? val.Replace("*", string.Empty) : "";
        }

        public static StringCollection GetTeamsForFinals(Worksheet worksheet)
        {
            return new StringCollection
            {
                worksheet.Range["DW23", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DW24", Type.Missing].Value2.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForSemiFinals(Worksheet worksheet)
        {
            return new StringCollection
            {
                worksheet.Range["DQ16", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DQ17", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DQ32", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DQ33", Type.Missing].Value2.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForQuarterFinals(Worksheet worksheet)
        {
            return new StringCollection
            {
                worksheet.Range["DK12", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DK20", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DK28", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DK36", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DK13", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DK21", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DK29", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DK37", Type.Missing].Value2.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForEightFinal(Worksheet worksheet)
        {
            if (!Tournament.IsGroupStageFinished(worksheet))
                return new StringCollection();

            return new StringCollection
            {
                worksheet.Range["DE10", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE14", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE18", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE22", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE26", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE30", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE34", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE38", Type.Missing].Value2.ToString().Replace("*", string.Empty),

                worksheet.Range["DE11", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE15", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE19", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE23", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE27", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE31", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE35", Type.Missing].Value2.ToString().Replace("*", string.Empty),
                worksheet.Range["DE39", Type.Missing].Value2.ToString().Replace("*", string.Empty)
            };
        }
    }
}
