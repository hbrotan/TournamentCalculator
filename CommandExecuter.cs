using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using TournamentCalculator.ExcelReaders;

namespace TournamentCalculator
{
    /// <summary>
    /// Summary description for CommandExecuter.
    /// </summary>
    public class CommandExecuter
    {
        private const string FILE_PREFIX = "EM2016";

        [STAThread]
        static void Main()
        {
            try
            {
                // create new command executer instance
                ExcelService.KillAllExcelProcesses();
                new CommandExecuter();
                Calculate();
                const string result = "Results Created. Press any key";
                Console.WriteLine(result);
                Console.ReadKey();
            }
            catch (Exception e)
            {
                ExcelService.KillAllExcelProcesses();

                const string result = "Error Occured. Press any key";
                Console.WriteLine(result);
                Console.Out.Write(e.Message);
                Console.ReadKey();
            }
        }

        private static void Calculate()
        {
            string fasitFile = ConfigurationManager.AppSettings["Fasit"];
            string sourceDirctory = ConfigurationManager.AppSettings["Source"];

            var oldCi = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            var resultWorksheet = ExcelService.GetResultWorksheet(fasitFile);

            var excel = new Application { Visible = false, UserControl = false };
            Worksheet correctResultsWorksheet = ExcelService.GetWorksheet(excel, resultWorksheet);
            StringCollection tablePosistions = GroupStage.GetTablePositions();

            // Fasit for sluttspill
            var results = GetResultsFromWorksheet(correctResultsWorksheet);
            
            // Regner ut poengsummene
            var scoresForAllUsers = new Dictionary<string, int>();
            foreach (var participant in Directory.GetFiles(sourceDirctory, "*.xlsx*"))
                AddParticipantScore(participant, excel, correctResultsWorksheet, tablePosistions, results, sourceDirctory, scoresForAllUsers);

            ResultFile.Create(scoresForAllUsers);

            ExcelService.Cleanup(excel);

            // reset old culture info
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCi;
        }

        private static Results GetResultsFromWorksheet(Worksheet correctResultsWorksheet)
        {
            return new Results
            {
                TeamsInEightFinal = TeamPlacementReader.GetTeamsForEightFinal(correctResultsWorksheet),
                TeamsInQuarterFinal = TeamPlacementReader.GetTeamsForQuarterFinals(correctResultsWorksheet),
                TeamsInSemiFinal = TeamPlacementReader.GetTeamsForSemiFinals(correctResultsWorksheet),
                TeamsInFinal = TeamPlacementReader.GetTeamsForFinals(correctResultsWorksheet),
                Winner = TeamPlacementReader.GetWinner(correctResultsWorksheet)
            };
        }

        private static void AddParticipantScore(string file, Application excel, Worksheet correctResultsWorksheet, StringCollection tablePosistions, Results results, string sourceDirctory, Dictionary<string, int> scoresForAllUsers)
        {
            var filename = Path.GetFileName(file);
            if (filename == null || !filename.StartsWith(FILE_PREFIX))
                return;

            Console.WriteLine("Processing {0}", file);

            Worksheet worksheet = ExcelService.GetWorksheet(excel, file);
            IEnumerable<int> matchesInGroupStage = GroupStage.GetMatches();
            int score = 0;

            // innledende kamper
            foreach (var i in matchesInGroupStage)
            {
                var r = correctResultsWorksheet.Range["F" + i.ToString(CultureInfo.InvariantCulture), Type.Missing];
                if (r.Value2 == null)
                    continue;

                var fasitHome = correctResultsWorksheet.Range["F" + i.ToString(CultureInfo.InvariantCulture), Type.Missing].Value2.ToString();
                var fasitAway = correctResultsWorksheet.Range["G" + i.ToString(CultureInfo.InvariantCulture), Type.Missing].Value2.ToString();
                var home = worksheet.Range["F" + i.ToString(CultureInfo.InvariantCulture), Type.Missing].Value2.ToString();
                var away = worksheet.Range["G" + i.ToString(CultureInfo.InvariantCulture), Type.Missing].Value2.ToString();

                if (GetHub(fasitHome, fasitAway).Equals(GetHub(home, away))) score += 2;
                if (fasitHome.Equals(home) && fasitAway.Equals(away)) score += 2;
            }

            // The table postitions, only if all matches are played                
            if (Tournament.IsGroupStageFinished(worksheet))
            {
                foreach (var tablePos in tablePosistions)
                {
                    var fasitPos = correctResultsWorksheet.Range[tablePos, Type.Missing].Value2.ToString();
                    var pos = worksheet.Range[tablePos, Type.Missing].Value2.ToString();
                    if (fasitPos.Equals(pos))
                        PointCalculator.AddScoreToGroupStageMatch(ref score, pos);
                }

                // The 1/8 finals
                var eight = TeamPlacementReader.GetTeamsForEightFinal(worksheet);
                foreach (var eightfinalists in results.TeamsInEightFinal.Cast<string>().Where(eight.Contains))
                    PointCalculator.AddScoreForEightFinals(ref score, eightfinalists);

                // The quarterfinals
                var quarter = TeamPlacementReader.GetTeamsForQuarterFinals(worksheet);
                foreach (var quarterfinalist in results.TeamsInQuarterFinal.Cast<string>().Where(quarter.Contains))
                    PointCalculator.AddScoreForQuarterfinals(ref score, quarterfinalist);

                // The semifinals
                var semis = TeamPlacementReader.GetTeamsForSemiFinals(worksheet);
                foreach (var semifinalist in results.TeamsInSemiFinal.Cast<string>().Where(semis.Contains))
                    PointCalculator.AddScoreForSemifinals(ref score, semifinalist);

                // The final
                var final = TeamPlacementReader.GetTeamsForFinals(worksheet);
                foreach (var finalist in results.TeamsInFinal.Cast<string>().Where(final.Contains))
                    PointCalculator.AddScoreForTeamInFinals(ref score, finalist);

                // The winner
                if (Tournament.IsWinnerDecided(worksheet))
                    PointCalculator.AddScoreForWinner(worksheet, results, ref score);
            }

            var name = file.Replace(sourceDirctory, "").Replace(FILE_PREFIX, "").Replace("_", " ").Replace(".xlsx", "").Replace("\\", "").Trim();

            scoresForAllUsers.Add(name, score);
        }

        /// <summary>
        /// Gets the HUB.
        /// </summary>
        /// <param name="home">The home.</param>
        /// <param name="away">The away.</param>
        /// <returns></returns>
        private static string GetHub(string home, string away)
        {
            if (Convert.ToInt32(home) > Convert.ToInt32(away)) return "H";
            if (Convert.ToInt32(home) == Convert.ToInt32(away)) return "U";
            return "B";
        }
    }
}
