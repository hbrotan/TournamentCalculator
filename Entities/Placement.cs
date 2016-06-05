namespace TournamentCalculator.Entities
{
    public class Placement
    {
        public int Rank { get; set; }
        public string Name { get; set; }
        public int Points { get; set; }
        public int? Trend { get; set; }
        public int? PointDifferenceFromYesterday { get; set; }
    }
}
