using System;
using System.Collections.Specialized;
using OfficeOpenXml;

namespace TournamentCalculator.ExcelReaders
{
    public class TeamPlacementReader
    {
        public static string GetWinner(ExcelWorksheet worksheet)
        {
            var val = Convert.ToString(worksheet.Cells["DT41"].Value);
            return !String.IsNullOrEmpty(val) ? val.Replace("*", string.Empty) : "";
        }

        public static StringCollection GetTeamsForFinals(ExcelWorksheet worksheet)
        {
            return new StringCollection
            {
                worksheet.Cells["DW23"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DW24"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForSemiFinals(ExcelWorksheet worksheet)
        {
            return new StringCollection
            {
                worksheet.Cells["DQ16"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DQ17"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DQ32"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DQ33"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForQuarterFinals(ExcelWorksheet worksheet)
        {
            return new StringCollection
            {
                worksheet.Cells["DK12"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DK20"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DK28"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DK36"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DK13"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DK21"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DK29"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DK37"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForEightFinal(ExcelWorksheet worksheet)
        {
            if (!Tournament.IsGroupStageFinished(worksheet))
                return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["DE10"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE14"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE18"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE22"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE26"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE30"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE34"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE38"].Value.ToString().Replace("*", string.Empty),

                worksheet.Cells["DE11"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE15"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE19"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE23"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE27"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE31"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE35"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["DE39"].Value.ToString().Replace("*", string.Empty)
            };
        }
    }
}