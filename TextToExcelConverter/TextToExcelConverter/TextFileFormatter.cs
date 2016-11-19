using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    class TextFileFormatter
    {
        private Dictionary<string, int> lastRowColumnFoundIn = new Dictionary<string, int>();
        private Dictionary<string, int> timesSeenInRow = new Dictionary<string, int>();

        InstanceConverter ic;
        public void formatText(ref int rowNumber, string line, InstanceConverter ic)
        {
            this.ic = ic;
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

        private string getColumnName(string columnName, int rowNumber)
        {
            if (lastRowColumnFoundIn.ContainsKey(columnName)) {
                if (ic.columns.ContainsKey(columnName) && lastRowColumnFoundIn[columnName] == rowNumber) {
                    if (!timesSeenInRow.ContainsKey(columnName)) {
                        timesSeenInRow.Add(columnName, 1);
                    }
                    timesSeenInRow[columnName]++;
                    columnName += "_" + timesSeenInRow[columnName];
                }
            }

            return columnName;
        }

        private void AddValueToRow(string columnName, int rowNum, string columnValue)
        {
            CreateColumn(columnName, rowNum);
            ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[rowNum, ic.columns[columnName]];
            ic.rng.Value = columnValue;
        }

        private void CreateColumn(string columnName, int rowNum)
        {
            if (!ic.columns.ContainsKey(columnName)) {
                ic.columns.Add(columnName, ic.columns.Count + 1);
                ic.ws.Cells[1, ic.columns.Count] = columnName;
                lastRowColumnFoundIn.Add(columnName, rowNum);
            }
        }
    }
}
