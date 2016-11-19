using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    class Statistics
    {
        Dictionary<string, int> lastRowColumnFoundIn = new Dictionary<string, int>();
        Dictionary<string, int> timesSeenInRow = new Dictionary<string, int>();

        Dictionary<string, int> fieldSeenTotal = new Dictionary<string, int>();
        Dictionary<string, int> fieldNullCount = new Dictionary<string, int>();

        public void GenerateStatistics(ref int rowNumber, string line, InstanceConverter ic)
        {
            if (line.ToLower().Contains("(")) {
                rowNumber++;
                timesSeenInRow.Clear();
                return;
            }
            if (line.Contains("=")) {
                Count(rowNumber, line, ic);
            }
        }

        public void OutputToExcel(InstanceConverter ic)
        {
            int i = 1;
            ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 1];
            ic.rng.Value = "Parameter Name";
            ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 2];
            ic.rng.Value = "# of times Parameter called";
            ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 3];
            ic.rng.Value = "# of times Parameter null or empty";
            ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 4];
            ic.rng.Value = "Usage Percentage";
            foreach (string key in fieldSeenTotal.Keys) {
                i++;
                ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 1];
                ic.rng.Value = key + ":";
                ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 2];
                ic.rng.Value = fieldSeenTotal[key];
                ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 3];
                ic.rng.Value = fieldNullCount[key];
                ic.rng = (Excel.Range)ic.wb.ActiveSheet.Cells[i, 4];
                double totalseen = fieldSeenTotal[key];
                double nullcount = fieldNullCount[key];
                ic.rng.Value = "%" + (int)(((totalseen - nullcount) / totalseen) * 100);
            }
        }

        private void Count(int rowNumber, string line, InstanceConverter ic)
        {
            string field = line.Split('=')[0].Trim();
            string value = line.Split('=')[1].Trim();
            field = getColumnName(field, rowNumber, ic);
            AddField(field, rowNumber);
            
            fieldSeenTotal[field]++;
            if (value == null || value.Equals("") || value.Equals("null")) {
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
            if (!lastRowColumnFoundIn.ContainsKey(field)) {
                lastRowColumnFoundIn.Add(field, rowNum);
            } else {
                lastRowColumnFoundIn[field] = rowNum;
            }
        }

        private string getColumnName(string fieldName, int rowNumber, InstanceConverter ic)
        {
            if (lastRowColumnFoundIn.ContainsKey(fieldName)) {
                if (ic.columns.ContainsKey(fieldName) && lastRowColumnFoundIn[fieldName] == rowNumber) {
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
