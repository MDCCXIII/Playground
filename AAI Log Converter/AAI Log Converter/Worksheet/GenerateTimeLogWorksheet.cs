using System.Collections.Generic;

namespace AAI_Log_Converter
{
    class GenerateTimeLogWorksheet
    {
        private Dictionary<string, int> columns = new Dictionary<string, int>();
        private FileConverter fileConverter;

        private const string wsName = "Time Log";

        public void GenerateTimeLog(int rowNumber, string line, FileConverter fileConverter)
        {
            this.fileConverter = fileConverter;
            if (line.StartsWith("),")) {
                string temp = line.Replace('(', ' ').Replace(')', ' ').Replace(',', ' ').Trim();
                string date = temp.Split()[0];
                string time = temp.Split()[1];
            }
        }

        private void AddValueToRow(string columnName, int rowNum, string columnValue)
        {
            CreateColumn(columnName, rowNum);
            Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(rowNum, columns[columnName], columnValue, wsName, false);
        }

        private void CreateColumn(string columnName, int rowNum)
        {
            if (!columns.ContainsKey(columnName)) {
                columns.Add(columnName, columns.Count + 1);
                Program.ExcelWriters[fileConverter.FileName].InsertValueIntoCell(1, columns[columnName], columnName, wsName, false);
            }
        }
    }
}
