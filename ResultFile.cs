using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using Newtonsoft.Json;

namespace TournamentCalculator
{
    public class ResultFile
    {
        public static void Create(Dictionary<string, int> scoresForAllUsers)
        {
            string resultFile = string.Format(ConfigurationManager.AppSettings["Result"], DateTime.Now.ToString("dd_MM_yyyy"));

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

            string json = JsonConvert.SerializeObject(scores.ToArray());

            System.IO.File.WriteAllText(resultFile, json);
        }

        private static int GetRank(KeyValuePair<string, int> entry, Placement prevPlacement, int currentRank)
        {
            if (prevPlacement == null)
                return 1;

            return entry.Value == prevPlacement.Points ? prevPlacement.Rank : currentRank;
        }

        private class Placement
        {
            public int Rank { get; set; }
            public string Name { get; set; }
            public int Points { get; set; }
        }
    }
}
