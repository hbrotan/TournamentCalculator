using System;
using System.Collections.Specialized;
using OfficeOpenXml;

namespace TournamentCalculator.ExcelReaders
{
    public class TeamPlacementReader
    {
        public static string GetWinner(ExcelWorksheet worksheet)
        {
            var val = Convert.ToString(worksheet.Cells["FO41"].Value);
            return !String.IsNullOrEmpty(val) ? val.Replace("*", string.Empty) : "";
        }

        //public static string GetBronzeWinner(ExcelWorksheet worksheet)
        //{
        //    var val = Convert.ToString(worksheet.Cells["BO41"].Value);
        //    return !String.IsNullOrEmpty(val) ? val.Replace("*", string.Empty) : "";
        //}

        public static StringCollection GetTeamsForFinals(ExcelWorksheet worksheet)
        {
            //if (!Tournament.IsSemiFinalsFinished(worksheet))
            //    return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["FR23"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FR24"].Value.ToString().Replace("*", string.Empty)
            };
        }

        //public static StringCollection GetTeamsForBronzeFinals(ExcelWorksheet worksheet)
        //{
        //    //if (!Tournament.IsSemiFinalsFinished(worksheet))
        //    //    return new StringCollection();

        //    return new StringCollection
        //    {
        //        worksheet.Cells["BR35"].Value.ToString().Replace("*", string.Empty),
        //        worksheet.Cells["BR36"].Value.ToString().Replace("*", string.Empty)
        //    };
        //}

        public static StringCollection GetTeamsForSemiFinals(ExcelWorksheet worksheet)
        {
            //if (!Tournament.IsQuarterFinalsFinished(worksheet))
            //    return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["FK16"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FK17"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FK32"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FK33"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForQuarterFinals(ExcelWorksheet worksheet)
        {
            //if (!Tournament.IsEightFinalsFinished(worksheet))
            //    return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["FD12"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FD20"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FD28"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FD36"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FD13"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FD21"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FD29"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["FD37"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForEightFinal(ExcelWorksheet worksheet)
        {
            if (!Tournament.IsGroupStageFinished(worksheet))
                return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["EW10"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW14"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW18"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW22"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW26"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW30"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW34"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW38"].Value.ToString().Replace("*", string.Empty),

                worksheet.Cells["EW11"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW15"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW19"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW23"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW27"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW31"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW35"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["EW39"].Value.ToString().Replace("*", string.Empty)
            };
        }
    }
}