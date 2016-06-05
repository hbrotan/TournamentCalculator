using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace TournamentCalculator.Entities
{
    public class ResultFile
    {
        public static string Create(Dictionary<string, int> scoresForAllUsers)
        {
            string resultFilePath = string.Format(ConfigurationManager.AppSettings["Result"], DateTime.Now.ToString("dd_MM_yyyy"));

            var scoresOrdered = scoresForAllUsers
                .OrderByDescending(x => x.Value)
                .ThenBy(x => x.Key);
            
            Placement previousEntry = null;
            int currentRank = 1;
            var scores = new List<Placement>();
            foreach (var entry in scoresOrdered)
            {
                var placement = new Placement
                {
                    Rank = GetRank(entry, previousEntry, currentRank),
                    Name = entry.Key,
                    Points = entry.Value
                };
                scores.Add(placement);
                previousEntry = placement;
                currentRank++;
            }

            var yesterdaysPlacements = GetScoresForYesterday();
            if(yesterdaysPlacements != null)
                scores = AddTrendAndPointsIncrease(scores, yesterdaysPlacements);

            string json = JsonConvert.SerializeObject(scores.ToArray());

            File.WriteAllText(resultFilePath, json);

            return json;
        }

        private static List<Placement> AddTrendAndPointsIncrease(List<Placement> placements, List<Placement> scoresForYesterday)
        {
            foreach (var placement in placements)
            {
                placement.Trend = placement.Rank - scoresForYesterday.Single(x => x.Name == placement.Name).Rank;
                placement.PointDifferenceFromYesterday = placement.Points - scoresForYesterday.Single(x => x.Name == placement.Name).Points;
            }
            return placements;
        }

        private static List<Placement> GetScoresForYesterday()
        {
            var filename = String.Empty;
            var lastDayWithResults = DateTime.Now.AddDays(-1);

            while (lastDayWithResults >= DateTime.Now.AddDays(-5))
            {
                filename = string.Format(ConfigurationManager.AppSettings["Result"], lastDayWithResults.ToString("dd_MM_yyyy"));

                if(File.Exists(filename))
                    break;

                lastDayWithResults = lastDayWithResults.AddDays(-1);
            }
            
            if (!File.Exists(filename))
                return null;

            var fileForYesterday = File.ReadAllText(filename);
            return JsonConvert.DeserializeObject<List<Placement>>(fileForYesterday);            
        }

        private static int GetRank(KeyValuePair<string, int> entry, Placement prevPlacement, int currentRank)
        {
            if (prevPlacement == null)
                return 1;

            return entry.Value == prevPlacement.Points ? prevPlacement.Rank : currentRank;
        }
    }
}
