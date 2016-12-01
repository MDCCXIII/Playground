using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AAI_Log_Converter
{
    class GenerateDataWorksheet
    {
        private Dictionary<string, int> lastRowColumnFoundIn = new Dictionary<string, int>();
        private Dictionary<string, int> timesSeenInRow = new Dictionary<string, int>();
        private Dictionary<string, int> columns = new Dictionary<string, int>();
        private Dictionary<int, int> RowNumbers = new Dictionary<int, int>();
        private FileConverter fileConverter;
        private static Mutex mut = new Mutex();
        private const string wsName = "Data";
        
        private int GetRow(int rowNumber, out int RowNumber)
        {
            if (!RowNumbers.ContainsKey(rowNumber)) {
                RowNumbers.Add(rowNumber, 0);
            }
            RowNumbers[rowNumber] = Program.ExcelWriters[fileConverter.FileName].GetNextRowNumber(wsName);
            RowNumber = RowNumbers[rowNumber];
            return RowNumbers[rowNumber];
        }

        public void formatText(int rowNumber, string line, FileConverter fileConverter)
        {
            this.fileConverter = fileConverter;
            mut.WaitOne();
            CreateColumn("Id", GetRow(rowNumber, out rowNumber));
            CreateColumn("TimeStamp", rowNumber);
            CreateColumn("Partner", rowNumber);
            mut.ReleaseMutex();
            if (line.ToLower().Contains("(")) {
                timesSeenInRow.Clear();
                return;
            }
            if (line.Contains("=")) {
                string columnName = line.Split('=')[0];
                string columnValue = line.Split('=')[1];
                columnName = getColumnName(columnName, rowNumber);
                CreateColumn(columnName, rowNumber);
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(rowNumber, columns[columnName], columnValue, wsName, false);
                return;
            }
            if (line.StartsWith("),")) {
                string temp = line.Replace('(', ' ').Replace(')', ' ').Replace(',', ' ').Trim();
                string date = temp.Split()[0];
                string time = temp.Split()[1];
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(rowNumber, columns["Id"], (rowNumber - 1).ToString(), wsName, false);
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(rowNumber, columns["TimeStamp"], date + " " + time, wsName, false);
                List<string> filestructure = fileConverter.filePath.Split('\\').ToList();
                string parentFolder = filestructure[filestructure.IndexOf(filestructure.Last()) - 1];
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(rowNumber, columns["Partner"], parentFolder, wsName, false);
                return;
            }
        }

        private string getColumnName(string columnName, int rowNumber)
        {
            if (lastRowColumnFoundIn.ContainsKey(columnName)) {
                if (columns.ContainsKey(columnName) && lastRowColumnFoundIn[columnName] == rowNumber) {
                    if (!timesSeenInRow.ContainsKey(columnName)) {
                        timesSeenInRow.Add(columnName, 1);
                    }
                    timesSeenInRow[columnName]++;
                    columnName += "_" + timesSeenInRow[columnName];
                }
            }
            return columnName;
        }

        private void CreateColumn(string columnName, int rowNum)
        {
            if (!columns.ContainsKey(columnName)) {
                columns.Add(columnName, columns.Count + 1);
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(1, columns[columnName], columnName, wsName, false);
                lastRowColumnFoundIn.Add(columnName, rowNum);
            } else {
                lastRowColumnFoundIn[columnName] = rowNum;
            }
        }
    }
}
