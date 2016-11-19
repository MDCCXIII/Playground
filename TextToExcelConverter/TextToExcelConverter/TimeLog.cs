using Excel = Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    class TimeLog
    {
        InstanceConverter ic;
        public void GenerateTimeLog(ref int rowNumber, string line, InstanceConverter ic)
        {
            this.ic = ic;
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
            }
        }
    }
}
