using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using TournamentCalculator.Entities;

namespace TournamentCalculator.ExcelReaders
{
    public class GroupStage
    {
        private const int NumberOfTeamsInGroup = 4;
        private const string ColumnTableStandings = "M";
        
        public static StringCollection GetTablePositions()
        {
            var groups = GetGroups();

            var tablePosistions = new StringCollection();
            foreach (var group in groups)
            {
                for (var i = 0; i < NumberOfTeamsInGroup; i++)
                {
                    var position = Convert.ToInt32(group.ExcelRow) + Convert.ToInt32(i);
                    tablePosistions.Add($"{group.ExcelColumn + position}");
                }
            }
            return tablePosistions;
        }


        public static IEnumerable<int> GetMatches()
        {            
            for (var i = 8; i <= 45; i++)
                yield return i;
        }

        private static IEnumerable<Group> GetGroups()
        {
            var groups = new List<Group>
            {
                //Hardkodet
                new Group {Letter = "A", ExcelColumn = ColumnTableStandings, ExcelRow = 9},
                new Group {Letter = "B", ExcelColumn = ColumnTableStandings, ExcelRow = 15},
                new Group {Letter = "C", ExcelColumn = ColumnTableStandings, ExcelRow = 21},
                new Group {Letter = "D", ExcelColumn = ColumnTableStandings, ExcelRow = 27},
                new Group {Letter = "E", ExcelColumn = ColumnTableStandings, ExcelRow = 33},
                new Group {Letter = "F", ExcelColumn = ColumnTableStandings, ExcelRow = 39}
            };
            return groups;
        }
    }
}