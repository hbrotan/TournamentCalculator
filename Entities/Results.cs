using System.Collections.Specialized;

namespace TournamentCalculator.Entities
{
    public class Results
    {
         public StringCollection TeamsInEightFinal { get; set; }

         public StringCollection TeamsInQuarterFinal { get; set; }

         public StringCollection TeamsInSemiFinal { get; set; } 

         public StringCollection TeamsInFinal { get; set; }

         public string Winner { get; set; }
    }
}