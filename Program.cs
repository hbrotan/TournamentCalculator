using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using TournamentCalculator.Entities;
using TournamentCalculator.ExcelReaders;

namespace TournamentCalculator
{
    public class Program
    {
        private const string FilePrefix = "VM2018";

        /// <summary>
        /// Assumes the following directories in same directory as executable (or different root)
        /// 
        /// /Leagues
        ///     /<LeagueName>
        ///         /Tippeforslag
        ///         /Resultat
        ///     Fasit.xslx
        /// 
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        static void Main(string[] args)
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", true, true);

            var configuration = builder.Build();

            if (!Directory.Exists(Path.Combine(Directory.GetCurrentDirectory(), "Leagues")))
            {
                Console.WriteLine($"Directory Leagues not found");
                Console.ReadKey();
                return;
            }

            if (!File.Exists(Path.Combine(Directory.GetCurrentDirectory(), "Leagues", "Fasit.xlsx")))
            {
                Console.WriteLine($"Fasit.xlsx not found in Leagues folder");
                Console.ReadKey();
                return;
            }

            try
            {
                var sourceDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Leagues");

                foreach (var league in Directory.GetDirectories(sourceDirectory))
                {
                    Console.WriteLine($"Processing league {Path.GetFileName(league)}");
                    var results = Calculate(sourceDirectory, league);
                    UploadResults(configuration["Tournament:Upload"], results, Path.GetFileName(league));
                }
            }
            catch (Exception e)
            {
                const string result = "Error Occurred. Press any key";
                Console.WriteLine(result);
                Console.Write(e.ToString());
                Console.ReadKey();
            }
        }

        private static void UploadResults(string uploadPath, string results, string leagueName)
        {
            var client = new HttpClient();
            var response = client
                .PostAsync(string.Format(uploadPath, FilePrefix, leagueName),
                    new StringContent(results, Encoding.UTF8, "application/json")).Result;

            Console.WriteLine(response.Content.ReadAsStringAsync().Result);
        }

        private static string Calculate(string sourcePath, string leaguePath)
        {
            var fasitFile = Path.Combine(sourcePath, "Fasit.xlsx");
            var sourceDirctory = Path.Combine(leaguePath, "Tippeforslag");
            var resultsDirectory = Path.Combine(leaguePath, "Resultat");

            var currentCulture = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            var correctResultsWorksheet = ExcelService.ExcelService.GetWorksheet(fasitFile);
            var tablePosistions = GroupStage.GetTablePositions();

            // Fasit for sluttspill
            var results = GetResultsFromWorksheet(correctResultsWorksheet);

            // Regner ut poengsummene
            var scoresForAllUsers = new List<UserScore>();
            foreach (var participant in Directory.GetFiles(sourceDirctory, "*.xlsx*"))
                AddParticipantScore(participant, correctResultsWorksheet, tablePosistions, results, sourceDirctory, scoresForAllUsers);

            var json = ResultFile.Create(scoresForAllUsers, resultsDirectory);

            // reset old culture info
            Thread.CurrentThread.CurrentCulture = currentCulture;

            return json;
        }

        private static Results GetResultsFromWorksheet(ExcelWorksheet correctResultsWorksheet)
        {
            return new Results
            {
                TeamsInEightFinal = TeamPlacementReader.GetTeamsForEightFinal(correctResultsWorksheet),
                TeamsInQuarterFinal = TeamPlacementReader.GetTeamsForQuarterFinals(correctResultsWorksheet),
                TeamsInSemiFinal = TeamPlacementReader.GetTeamsForSemiFinals(correctResultsWorksheet),
                TeamsInBronzeFinal = TeamPlacementReader.GetTeamsForBronzeFinals(correctResultsWorksheet),
                TeamsInFinal = TeamPlacementReader.GetTeamsForFinals(correctResultsWorksheet),
                Winner = TeamPlacementReader.GetWinner(correctResultsWorksheet)
            };
        }

        private static void AddParticipantScore(string file, ExcelWorksheet correctResultsWorksheet, StringCollection tablePosistions, Results results, string sourceDirctory, List<UserScore> scoresForAllUsers)
        {
            var filename = Path.GetFileName(file);
            if (filename == null || !filename.StartsWith(FilePrefix))
                return;

            Console.WriteLine("Processing {0}", file);

            var worksheet = ExcelService.ExcelService.GetWorksheet(file);

            if (worksheet.Cells["A1"].Value.ToString() != "Verdensmesterskapet i fotball 2018")
            {
                Console.WriteLine($"Language not Norwegian for: {filename}");
                Console.WriteLine($"Excel sheets will be omitted. Press enter to continue processing the next sheet");
                Console.ReadLine();
                return;
            }

            var matchesInGroupStage = GroupStage.GetMatches();
            var score = 0;

            // innledende kamper
            foreach (var i in matchesInGroupStage)
            {
                var r = correctResultsWorksheet.Cells["F" + i.ToString(CultureInfo.InvariantCulture)];
                if (r.Value == null)
                    continue;

                if (worksheet.Cells["F" + i.ToString(CultureInfo.InvariantCulture)].Value == null || worksheet.Cells["G" + i.ToString(CultureInfo.InvariantCulture)].Value == null)
                {
                    Console.WriteLine($"Group stage not correctly filled out for: {filename}");
                    Console.WriteLine("Excel sheet will be omitted. Press enter to continue processing the next sheet");

                    Console.ReadLine();
                    return;
                }

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
                if (worksheet.Cells["BA10"].Value == null)
                {
                    Console.WriteLine($"Knockout stage not correctly filled out for: {filename}");
                    Console.WriteLine($"Excel sheets will be omitted. Press enter to continue processing the next sheet");
                    Console.ReadLine();
                    return;
                }

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

                // The bronze final
                var bronzeFinal = TeamPlacementReader.GetTeamsForBronzeFinals(worksheet);
                foreach (var finalist in results.TeamsInBronzeFinal.Cast<string>().Where(bronzeFinal.Contains))
                    PointCalculator.AddScoreForTeamInBronzeFinals(ref score, finalist);

                // The final
                var final = TeamPlacementReader.GetTeamsForFinals(worksheet);
                foreach (var finalist in results.TeamsInFinal.Cast<string>().Where(final.Contains))
                    PointCalculator.AddScoreForTeamInFinals(ref score, finalist);

                // The bronze final
                if (Tournament.IsBronzeWinnerDecided(correctResultsWorksheet))
                {
                    var fasitHome = correctResultsWorksheet.Cells["BS35"].Value.ToString();
                    var fasitAway = correctResultsWorksheet.Cells["BS36"].Value.ToString();

                    if (worksheet.Cells["BS35"].Value == null || worksheet.Cells["BS36"].Value == null)
                    {
                        Console.WriteLine($"Bronze final not correctly filled out for: {filename}");
                        Console.WriteLine("Excel sheet will be omitted. Press enter to continue processing the next sheet");
                        Console.ReadLine();
                        return;
                    }

                    var home = worksheet.Cells["BS35"].Value.ToString();
                    var away = worksheet.Cells["BS36"].Value.ToString();

                    if (GetHub(fasitHome, fasitAway) == "H" && GetHub(home, away) == "H" && bronzeFinal[0] == results.TeamsInBronzeFinal[0])
                        PointCalculator.AddScoreForBronzeWinner(ref score, results.TeamsInBronzeFinal[0]);

                    if (GetHub(fasitHome, fasitAway) == "B" && GetHub(home, away) == "B" && bronzeFinal[1] == results.TeamsInBronzeFinal[1])
                        PointCalculator.AddScoreForBronzeWinner(ref score, results.TeamsInBronzeFinal[1]);
                }

                // The winner
                if (Tournament.IsWinnerDecided(worksheet))
                    PointCalculator.AddScoreForWinner(worksheet, results, ref score);
            }

            var name = file.Replace(sourceDirctory, "").Replace(FilePrefix, "").Replace("_", " ").Replace(".xlsx", "").Replace("\\", "").Trim();

            scoresForAllUsers.Add(new UserScore{ Name = name, Points = score, Winner = TeamPlacementReader.GetWinner(worksheet)});
        }

        private static bool HasValidLanguage(ExcelWorksheet worksheet, string fileName)
        {
            if (worksheet.Cells["O3"].Value.ToString() != "Language: Norwegian")
            {
                Console.WriteLine($"Language not Norwegian for: {fileName}");
                Console.WriteLine("Excel sheet will be omitted. Press enter to continue processing the next sheet");
                return false;
            }

            return true;
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
