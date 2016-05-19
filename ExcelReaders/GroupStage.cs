using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using TournamentCalculator.Entities;

namespace TournamentCalculator.ExcelReaders
{
    public class GroupStage
    {
        private const int NUMBER_OF_TEAMS_IN_GROUP = 4;

        private const string COLUMN_TABLE_STANDINGS = "M";
        
        public static StringCollection GetTablePositions()
        {
            var groups = GetGroups();

            var tablePosistions = new StringCollection();
            foreach (var group in groups)
            {
                for (var i = 0; i < NUMBER_OF_TEAMS_IN_GROUP; i++)
                {
                    var pos = Convert.ToInt32(@group.ExcelRow) + Convert.ToInt32(i);
                    tablePosistions.Add(String.Format("{0}", @group.ExcelColumn + pos));
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
                new Group {Letter = "A", ExcelColumn = COLUMN_TABLE_STANDINGS, ExcelRow = 9},
                new Group {Letter = "B", ExcelColumn = COLUMN_TABLE_STANDINGS, ExcelRow = 15},
                new Group {Letter = "C", ExcelColumn = COLUMN_TABLE_STANDINGS, ExcelRow = 21},
                new Group {Letter = "D", ExcelColumn = COLUMN_TABLE_STANDINGS, ExcelRow = 27},
                new Group {Letter = "E", ExcelColumn = COLUMN_TABLE_STANDINGS, ExcelRow = 33},
                new Group {Letter = "F", ExcelColumn = COLUMN_TABLE_STANDINGS, ExcelRow = 39}
            };
            return groups;
        }
    }
}
