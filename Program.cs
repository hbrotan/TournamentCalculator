using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using TournamentCalculator.Entities;
using TournamentCalculator.ExcelReaders;

namespace TournamentCalculator
{
    public class Program
    {
        private const string FilePrefix = "EM2016";

        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", true, true);

            var configuration = builder.Build();

            try
            {
                var results = Calculate(configuration);

                var client = new HttpClient();
                var response = client
                    .PostAsync(string.Format(
                            "https://tournament.azurewebsites.net/api/tournament/{0}/league/{1}/result",
                            FilePrefix, configuration["Tournament:Liganavn"]),
                        new StringContent(results, Encoding.UTF8, "application/json")).Result;

                Console.WriteLine(response.Content.ReadAsStringAsync().Result);
            }
            catch (Exception e)
            {
                const string result = "Error Occured. Press any key";
                Console.WriteLine(result);
                Console.Out.Write(e.Message);
                Console.ReadKey();
            }
        }

        private static string Calculate(IConfiguration configuration)
        {
            var fasitFile = configuration["Tournament:Fasit"];
            var sourceDirctory = configuration["Tournament:Source"];
            var resultFilePattern = configuration["Tournament:Result"];

            var oldCi = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            var resultWorksheet = ExcelService.ExcelService.GetResultWorksheet(fasitFile);
            var correctResultsWorksheet = ExcelService.ExcelService.GetWorksheet(resultWorksheet);
            var tablePosistions = GroupStage.GetTablePositions();

            // Fasit for sluttspill
            var results = GetResultsFromWorksheet(correctResultsWorksheet);
            
            // Regner ut poengsummene
            var scoresForAllUsers = new Dictionary<string, int>();
            foreach (var participant in Directory.GetFiles(sourceDirctory, "*.xlsx*"))
                AddParticipantScore(participant, correctResultsWorksheet, tablePosistions, results, sourceDirctory, scoresForAllUsers);

            var json = ResultFile.Create(scoresForAllUsers, resultFilePattern);

            // reset old culture info
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCi;

            return json;
        }

        private static Results GetResultsFromWorksheet(ExcelWorksheet correctResultsWorksheet)
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

        private static void AddParticipantScore(string file, ExcelWorksheet correctResultsWorksheet, StringCollection tablePosistions, Results results, string sourceDirctory, Dictionary<string, int> scoresForAllUsers)
        {
            var filename = Path.GetFileName(file);
            if (filename == null || !filename.StartsWith(FilePrefix))
                return;

            Console.WriteLine("Processing {0}", file);

            var worksheet = ExcelService.ExcelService.GetWorksheet(file);
            var matchesInGroupStage = GroupStage.GetMatches();
            var score = 0;

            // innledende kamper
            foreach (var i in matchesInGroupStage)
            {
                var r = correctResultsWorksheet.Cells["F" + i.ToString(CultureInfo.InvariantCulture)];
                if (r.Value == null)
                    continue;

                var fasitHome = correctResultsWorksheet.Cells["F" + i.ToString(CultureInfo.InvariantCulture)].Value.ToString();
                var fasitAway = correctResultsWorksheet.Cells["G" + i.ToString(CultureInfo.InvariantCulture)].Value.ToString();
                var home = worksheet.Cells["F" + i.ToString(CultureInfo.InvariantCulture)].Value.ToString();
                var away = worksheet.Cells["G" + i.ToString(CultureInfo.InvariantCulture)].Value.ToString();

                if (GetHub(fasitHome, fasitAway).Equals(GetHub(home, away)))
                    PointCalculator.AddScoreForCorrectOutcomeInGroupMatch(ref score);

                if (fasitHome.Equals(home) && fasitAway.Equals(away)) 
                    PointCalculator.AddScoreForCorrectResultInGroupMatch(ref score);
            }

            // The table postitions, only if all matches are played                
            if (Tournament.IsGroupStageFinished(correctResultsWorksheet))
            {
                foreach (var tablePos in tablePosistions)
                {
                    var fasitPos = correctResultsWorksheet.Cells[tablePos].Value.ToString();
                    var pos = worksheet.Cells[tablePos].Value.ToString();
                    if (fasitPos.Equals(pos))
                        PointCalculator.AddScoreForCorrectPlacementInGroup(ref score, pos);
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

            var name = file.Replace(sourceDirctory, "").Replace(FilePrefix, "").Replace("_", " ").Replace(".xlsx", "").Replace("\\", "").Trim();
            
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
