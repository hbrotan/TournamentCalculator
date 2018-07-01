using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace TournamentCalculator.Entities
{
    public class ResultFile
    {
        public static string Create(List<UserScore> scoresForAllUsers, string resultsDirectory)
        {
            var resultFilePath = $"{resultsDirectory}\\Resultat_{DateTime.Now:dd_MM_yyyy}.json";

            var scoresOrdered = scoresForAllUsers
                .OrderByDescending(x => x.Points)
                .ThenBy(x => x.Name);
            
            Placement previousEntry = null;
            var currentRank = 1;
            var scores = new List<Placement>();
            foreach (var entry in scoresOrdered)
            {
                var placement = new Placement
                {
                    Rank = GetRank(entry, previousEntry, currentRank),
                    Name = entry.Name,
                    Points = entry.Points,
                    Winner = entry.Winner
                };
                scores.Add(placement);
                previousEntry = placement;
                currentRank++;
            }

            var yesterdaysPlacements = GetScoresForYesterday(resultsDirectory);
            if(yesterdaysPlacements != null)
                scores = AddTrendAndPointsIncrease(scores, yesterdaysPlacements);

            var json = JsonConvert.SerializeObject(scores.ToArray());

            File.WriteAllText(resultFilePath, json);

            return json;
        }

        private static List<Placement> AddTrendAndPointsIncrease(List<Placement> placements, List<Placement> scoresForYesterday)
        {
            foreach (var placement in placements)
            {
                var score = scoresForYesterday.SingleOrDefault(x => x.Name == placement.Name);
                if (score != null)
                {
                    placement.Trend = placement.Rank - score.Rank;
                    placement.PointDifferenceFromYesterday = placement.Points - score.Points;
                }
            }
            return placements;
        }

        private static List<Placement> GetScoresForYesterday(string resultsDirectory)
        {
            var filename = String.Empty;
            var lastDayWithResults = DateTime.Now.AddDays(-1);

            while (lastDayWithResults >= DateTime.Now.AddDays(-5))
            {
                filename = $"{resultsDirectory}\\Resultat_{lastDayWithResults:dd_MM_yyyy}.json";

                if (File.Exists(filename))
                    break;

                lastDayWithResults = lastDayWithResults.AddDays(-1);
            }
            
            if (!File.Exists(filename))
                return null;

            var fileForYesterday = File.ReadAllText(filename);
            return JsonConvert.DeserializeObject<List<Placement>>(fileForYesterday);
        }

        private static int GetRank(UserScore entry, Placement prevPlacement, int currentRank)
        {
            if (prevPlacement == null)
                return 1;

            return entry.Points == prevPlacement.Points ? prevPlacement.Rank : currentRank;
        }
    }
}