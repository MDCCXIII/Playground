using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    class TextFileFormatter
    {
        public static Dictionary<string, int> lastRowColumnFoundIn = new Dictionary<string, int>();
        public static Dictionary<string, int> timesSeenInRow = new Dictionary<string, int>();
        public static void formatText(ref int rowNumber, string line)
        {
            CreateColumn("Id", rowNumber);
            CreateColumn("TimeStamp", rowNumber);
            if (line.ToLower().Contains("(")) {
                rowNumber++;
                timesSeenInRow.Clear();
                return;
            }
            if (line.Contains("=")) {
                string columnName = line.Split('=')[0];
                string columnValue = line.Split('=')[1];
                columnName = getColumnName(columnName, rowNumber);
                AddValueToRow(columnName, rowNumber, columnValue);
                return;
            }
            if (line.StartsWith("),")) {
                string temp = line.Replace('(', ' ').Replace(')', ' ').Replace(',', ' ').Trim();
                string date = temp.Split()[0];
                string time = temp.Split()[1];
                AddValueToRow("Id", rowNumber, (rowNumber - 1).ToString());
                AddValueToRow("TimeStamp", rowNumber, date + " " + time);
                return;
            }
        }

        private static string getColumnName(string columnName, int rowNumber)
        {
            if (lastRowColumnFoundIn.ContainsKey(columnName)) {
                if (Form1.columns.ContainsKey(columnName) && lastRowColumnFoundIn[columnName] == rowNumber) {
                    if (!timesSeenInRow.ContainsKey(columnName)) {
                        timesSeenInRow.Add(columnName, 1);
                    }
                    timesSeenInRow[columnName]++;
                    columnName += "_" + timesSeenInRow[columnName];
                }
            }

            return columnName;
        }

        private static void AddValueToRow(string columnName, int rowNum, string columnValue)
        {
            CreateColumn(columnName, rowNum);
            Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[rowNum, Form1.columns[columnName]];
            Form1.rng.Value = columnValue;
        }

        private static void CreateColumn(string columnName, int rowNum)
        {
            if (!Form1.columns.ContainsKey(columnName)) {
                Form1.columns.Add(columnName, Form1.columns.Count + 1);
                Form1.ws.Cells[1, Form1.columns.Count] = columnName;
                lastRowColumnFoundIn.Add(columnName, rowNum);
            }
        }
    }
}
