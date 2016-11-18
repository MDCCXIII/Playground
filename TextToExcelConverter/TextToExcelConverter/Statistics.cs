using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    class Statistics
    {
        public static Dictionary<string, int> lastRowColumnFoundIn = new Dictionary<string, int>();
        public static Dictionary<string, int> timesSeenInRow = new Dictionary<string, int>();

        public static Dictionary<string, int> fieldSeenTotal = new Dictionary<string, int>();
        public static Dictionary<string, int> fieldNullCount = new Dictionary<string, int>();

        public static void GenerateStatistics(ref int rowNumber, string line)
        {
            if (line.ToLower().Contains("(")) {
                rowNumber++;
                timesSeenInRow.Clear();
                return;
            }
            if (line.Contains("=")) {
                Count(rowNumber, line);
            }
        }

        public static void OutputToExcel()
        {
            int i = 1;
            Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 1];
            Form1.rng.Value = "Parameter Name";
            Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 2];
            Form1.rng.Value = "# of times Parameter called";
            Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 3];
            Form1.rng.Value = "# of times Parameter null or empty";
            Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 4];
            Form1.rng.Value = "Usage Percentage";
            foreach (string key in fieldSeenTotal.Keys) {
                i++;
                Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 1];
                Form1.rng.Value = key + ":";
                Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 2];
                Form1.rng.Value = fieldSeenTotal[key];
                Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 3];
                Form1.rng.Value = fieldNullCount[key];
                Form1.rng = (Excel.Range)Form1.wb.ActiveSheet.Cells[i, 4];
                double totalseen = fieldSeenTotal[key];
                double nullcount = fieldNullCount[key];
                Form1.rng.Value = "%" + (int)(((totalseen - nullcount) / totalseen) * 100);
            }
        }

        private static void Count(int rowNumber, string line)
        {
            string field = line.Split('=')[0].Trim();
            string value = line.Split('=')[1].Trim();
            field = getColumnName(field, rowNumber);
            AddField(field, rowNumber);
            
            fieldSeenTotal[field]++;
            if (value == null || value.Equals("") || value.Equals("null")) {
                fieldNullCount[field]++;
            }
        }

        private static void AddField(string field, int rowNum)
        {
            if (!fieldSeenTotal.ContainsKey(field)) {
                fieldSeenTotal.Add(field, 0);
            }
            if (!fieldNullCount.ContainsKey(field)) {
                fieldNullCount.Add(field, 0);
            }
            if (!lastRowColumnFoundIn.ContainsKey(field)) {
                lastRowColumnFoundIn.Add(field, rowNum);
            } else {
                lastRowColumnFoundIn[field] = rowNum;
            }
        }

        private static string getColumnName(string fieldName, int rowNumber)
        {
            if (lastRowColumnFoundIn.ContainsKey(fieldName)) {
                if (Form1.columns.ContainsKey(fieldName) && lastRowColumnFoundIn[fieldName] == rowNumber) {
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
