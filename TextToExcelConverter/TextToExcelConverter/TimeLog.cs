using Excel = Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    class TimeLog
    {
        public static void GenerateTimeLog(ref int rowNumber, string line)
        {
            if (line.StartsWith("),")) {
                rowNumber++;
                string temp = line.Replace('(', ' ').Replace(')', ' ').Replace(',', ' ').Trim();
                string date = temp.Split()[0];
                string time = temp.Split()[1];
                AddValueToRow("Id", rowNumber, (rowNumber - 1).ToString());
                AddValueToRow("TimeStamp", rowNumber, date + " " + time);
                return;
            }
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
            }
        }
    }
}
