using System;
using System.Collections.Specialized;
using OfficeOpenXml;

namespace TournamentCalculator.ExcelReaders
{
    public class TeamPlacementReader
    {
        public static string GetWinner(ExcelWorksheet worksheet)
        {
            var val = Convert.ToString(worksheet.Cells["BO41"].Value);
            return !String.IsNullOrEmpty(val) ? val.Replace("*", string.Empty) : "";
        }

        public static string GetBronzeWinner(ExcelWorksheet worksheet)
        {
            var val = Convert.ToString(worksheet.Cells["BO41"].Value);
            return !String.IsNullOrEmpty(val) ? val.Replace("*", string.Empty) : "";
        }

        public static StringCollection GetTeamsForFinals(ExcelWorksheet worksheet)
        {
            //if (!Tournament.IsSemiFinalsFinished(worksheet))
            //    return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["BR23"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BR24"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForBronzeFinals(ExcelWorksheet worksheet)
        {
            //if (!Tournament.IsSemiFinalsFinished(worksheet))
            //    return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["BR35"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BR36"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForSemiFinals(ExcelWorksheet worksheet)
        {
            //if (!Tournament.IsQuarterFinalsFinished(worksheet))
            //    return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["BL16"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BL17"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BL32"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BL33"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForQuarterFinals(ExcelWorksheet worksheet)
        {
            //if (!Tournament.IsEightFinalsFinished(worksheet))
            //    return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["BF12"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BF20"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BF28"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BF36"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BF13"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BF21"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BF29"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["BF37"].Value.ToString().Replace("*", string.Empty)
            };
        }

        public static StringCollection GetTeamsForEightFinal(ExcelWorksheet worksheet)
        {
            if (!Tournament.IsGroupStageFinished(worksheet))
                return new StringCollection();

            return new StringCollection
            {
                worksheet.Cells["AZ10"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ14"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ18"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ22"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ26"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ30"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ34"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ38"].Value.ToString().Replace("*", string.Empty),

                worksheet.Cells["AZ11"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ15"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ19"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ23"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ27"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ31"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ35"].Value.ToString().Replace("*", string.Empty),
                worksheet.Cells["AZ39"].Value.ToString().Replace("*", string.Empty)
            };
        }
    }
}