using System;
using System.Text;
using OfficeOpenXml;
using TournamentCalculator.Entities;

namespace TournamentCalculator.ExcelReaders
{
    public class PointCalculator
    {
        public static int AddScoreForWinner(ref int score, string winner)
        {
            Console.WriteLine("+16 for korrekt finalevinner : {0}", winner);
            score += 16;
            return score;
        }

        public static int AddScoreForBronzeWinner(ref int score, string bronzeWinner)
        {
            Console.WriteLine("+14 for korrekt bronsefinalevinner : {0}", bronzeWinner);
            score += 14;
            return score;
        }

        public static void AddScoreForCorrectResultInGroupMatch(ref int score)
        {
            score += 2;
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("+2 for gruppespillkamp : korrekt resultat");
        }

        public static void AddScoreForCorrectOutcomeInGroupMatch(ref int score)
        {
            score += 2;
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("+2 for gruppespillkamp : korrekt utfall");
        }

        public static void AddScoreForCorrectPlacementInGroup(ref int score, dynamic pos)
        {
            score += 2;
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("+2 for {0} på korrekt plass i gruppen", pos);
        }

        public static void AddScoreForEightFinals(ref int score, string eightfinalists)
        {
            score += 4;
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine("+4 for {0} videre til åttendelsfinale", eightfinalists);
        }

        public static void AddScoreForQuarterfinals(ref int score, string quarterfinalist)
        {
            score += 6;
            Console.WriteLine("+6 for {0} videre til kvartfinale", quarterfinalist);
        }

        public static void AddScoreForSemifinals(ref int score, string semifinalist)
        {
            score += 8;
            Console.WriteLine("+8 for {0} videre til semifinale", semifinalist);
        }

        public static void AddScoreForTeamInFinals(ref int score, string finalist)
        {
            score += 12;
            Console.WriteLine("+12 for {0} videre til finale", finalist);
        }

        public static void AddScoreForTeamInBronzeFinals(ref int score, string finalist)
        {
            score += 10;
            Console.WriteLine("+10 for {0} videre til bronsefinale", finalist);
        }

        public static void AddScoreForWinner(ExcelWorksheet worksheet, Results results, ref int score)
        {
            var winner = TeamPlacementReader.GetWinner(worksheet);

            if (winner.Equals(results.Winner))
                score = AddScoreForWinner(ref score, winner);
        }

        public static void AddScoreForBronzeWinner(ExcelWorksheet worksheet, Results results, ref int score)
        {
            var bronzeWinner = TeamPlacementReader.GetBronzeWinner(worksheet);

            if (bronzeWinner.Equals(results.BronzeWinner))
                score = AddScoreForBronzeWinner(ref score, bronzeWinner);
        }
    }
}