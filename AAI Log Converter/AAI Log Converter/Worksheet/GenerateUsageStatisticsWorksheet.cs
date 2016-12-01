using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AAI_Log_Converter
{
    class GenerateUsageStatisticsWorksheet
    {
        private Dictionary<string, int> lastRowColumnFoundIn = new Dictionary<string, int>();
        private Dictionary<string, int> timesSeenInRow = new Dictionary<string, int>();

        private Dictionary<string, int> fieldSeenTotal = new Dictionary<string, int>();
        private Dictionary<string, int> fieldNullCount = new Dictionary<string, int>();
        private Dictionary<string, int> fieldEmptyCount = new Dictionary<string, int>();

        private const string wsName = "Usage Statistics";
        
        public void GenerateStatistics(int rowNumber, string line)
        {
            if (line.ToLower().Contains("(")) {
                timesSeenInRow.Clear();
                return;
            }
            if (line.Contains("=")) {
                Count(rowNumber, line);
            }
        }

        public void OutputToExcel(FileConverter fileConverter)
        {
            int i = 1;
            Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 1, "Parameter Name", wsName, false);
            Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 2, "# of times parameter called", wsName, false);
            Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 3, "# of times parameter null", wsName, false);
            Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 4, "# of times parameter empty", wsName, false);
            Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 5, "Usage Percentage", wsName, false);

            foreach (string key in fieldSeenTotal.Keys) {
                i++;
                int oldval;
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 1, key + ":", wsName, true);
               
                oldval = Program.ExcelWriters[fileConverter.FileName].GetCellInteger(i, 2, wsName);
                double totalseen = fieldSeenTotal[key] + oldval;
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 2, (fieldSeenTotal[key] + oldval).ToString(), wsName, true);

                oldval = Program.ExcelWriters[fileConverter.FileName].GetCellInteger(i, 3, wsName);
                double nullcount = fieldNullCount[key] + oldval;
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 3, (fieldNullCount[key] + oldval).ToString(), wsName, true);

                oldval = Program.ExcelWriters[fileConverter.FileName].GetCellInteger(i, 4, wsName);
                double emptycount = fieldEmptyCount[key] + oldval;
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 4, (fieldEmptyCount[key] + oldval).ToString(), wsName, true);
                
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(i, 5, "%" + (int)(((totalseen - (nullcount + emptycount)) / totalseen) * 100), wsName, true);
            }
        }

        private void Count(int rowNumber, string line)
        {
            string field = line.Split('=')[0].Trim();
            string value = line.Split('=')[1].Trim();
            field = getColumnName(field, rowNumber);
            AddField(field, rowNumber);

            fieldSeenTotal[field]++;
            if (value == null || value.Equals("")) {
                fieldEmptyCount[field]++;
            }
            if (value != null && value.Equals("null")) {
                fieldNullCount[field]++;
            }
        }

        private void AddField(string field, int rowNum)
        {
            if (!fieldSeenTotal.ContainsKey(field)) {
                fieldSeenTotal.Add(field, 0);
            }
            if (!fieldNullCount.ContainsKey(field)) {
                fieldNullCount.Add(field, 0);
            }
            if (!fieldEmptyCount.ContainsKey(field)) {
                fieldEmptyCount.Add(field, 0);
            }
            if (!lastRowColumnFoundIn.ContainsKey(field)) {
                lastRowColumnFoundIn.Add(field, rowNum);
            } else {
                lastRowColumnFoundIn[field] = rowNum;
            }
        }

        private string getColumnName(string fieldName, int rowNumber)
        {
            if (lastRowColumnFoundIn.ContainsKey(fieldName)) {
                if (lastRowColumnFoundIn[fieldName] == rowNumber) {
                    if (!timesSeenInRow.ContainsKey(fieldName)) {
                        timesSeenInRow.Add(fieldName, 1);
                    }
                    timesSeenInRow[fieldName]++;
                    fieldName += "_" + timesSeenInRow[fieldName];
                }
            }

            return fieldName;
        }
    }
}
