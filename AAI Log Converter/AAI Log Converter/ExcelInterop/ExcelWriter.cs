using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace AAI_Log_Converter.ExcelInterop
{
    class ExcelWriter
    {
        #region Globals

        public Excel.Workbook workBook = null;
        private static Mutex mut = new Mutex();

        #endregion

        #region Constructors

        #endregion

        #region Accessors

        public void CreateOrActivatWorkSheet(string wsName)
        {
            ExistOrCreate(wsName);
        }

        public void InsertValueIntoCell(int rowNumber, int columnNumber, string value, string wsName, bool overwrite)
        {
            Lock();
            if (overwrite) {
                CreateOrActivatWorkSheet(wsName);
                workBook.Worksheets[wsName].Cells[rowNumber, columnNumber].Value = value;
            } else {
                if(workBook.Worksheets[wsName].Cells[rowNumber, columnNumber].Value == null) {
                    CreateOrActivatWorkSheet(wsName);
                    workBook.Worksheets[wsName].Cells[rowNumber, columnNumber].Value = value;
                }
            }
            Release();
        }

        public void SaveExcelFile(string fileName)
        {
            int attempt = 0;
            bool failed = false;
                do {
                        try {
                            attempt++;
                            workBook.SaveAs(Directory.GetCurrentDirectory() + "\\" + fileName + ".xlsx");
                            failed = false;
                        }
                        catch (COMException) {
                            failed = true;
                            //controlSet.setInfoLabelText("Attempt to save file failed. Attempt " + attempt);
                            //TODO: Add Logging
                        }
                        Thread.Sleep(10);
                    //}
                } while (failed && attempt > 3);
                if (failed) {
                    //controlSet.setInfoLabelText("Failed to save the converted file.");
                    //TODO: Add Logging
                } else {
                   // controlSet.setInfoLabelText("Conversion Complete.");
                    //TODO: Add Logging
                }
            //}
        }

        public void Cleanup()
        {
            try {
                workBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(workBook);
                workBook = null;
            }
            catch (Exception) {
                //TODO: Add Logging
            }
            finally {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        Dictionary<string, int> rowIndices = new Dictionary<string, int>();
        private void AddRowIndices(string worksheetName)
        {
            if (!rowIndices.ContainsKey(worksheetName)) {
                rowIndices.Add(worksheetName, 0);
            }
        }

        public int GetNextRowNumber(string tableName)
        {
            Lock();
            CreateOrActivatWorkSheet(tableName);
            AddRowIndices(tableName);
            Excel.Worksheet ws = workBook.Worksheets[tableName];
            Excel.Range range = (Excel.Range)ws.Cells[ws.Rows.Count, 1];
            rowIndices[tableName] = (int)range.get_End(Excel.XlDirection.xlUp).Row + 1;
            if(rowIndices[tableName] == 1) {
                rowIndices[tableName]++;
            }
            Release();

            return rowIndices[tableName];
        }

        public int GetCellInteger(int rowNumber, int columnNumber, string sheetName)
        {
            Lock();
            int result = 0;
            CreateOrActivatWorkSheet(sheetName);
            Excel.Worksheet ws = workBook.Worksheets[sheetName];
            string cellValue = ws.Cells[rowNumber, columnNumber].Text.ToString();
            if(!Int32.TryParse(cellValue, out result)) {
                result = 0;
            }
            Release();
            return result;
        }

        #endregion

        #region Internals

        
        private static void Lock()
        {
            mut.WaitOne();
        }

        private static void Release()
        {
            mut.ReleaseMutex();
        }

        private bool ExistOrCreate(string wsName)
        {
            try {
                Excel.Worksheet t = workBook.Worksheets[wsName];
                //controlSet.setInfoLabelText("Set the " + wsName + " worksheet to the active sheet.");
            }
            catch (COMException comEx) {
                //TODO: Add Logging
                try {
                    var xlSheets = workBook.Sheets as Excel.Sheets;
                    var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                    xlNewSheet.Name = wsName;
                    //controlSet.setInfoLabelText("Created a new worksheet: " + wsName + ".");
                }
                catch (Exception ex) {
                    //TODO: Add Logging
                    //controlSet.setInfoLabelText("Failed to Create or activate a worksheet.");
                }

            }
            return true;
        }

        public void OpenOrCreateNewWorkBook(string excelFilePath)
        {
            try {
                workBook = new Excel.Application().Workbooks._Open(excelFilePath, System.Reflection.Missing.Value, false); // open the existing excel file
                //controlSet.setInfoLabelText("Opened existing workbook.");
                //TODO: Add Logging

            }
            catch (Exception) {
                workBook = new Excel.Application().Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                //controlSet.setInfoLabelText("Created new workbook.");
            }

        }

        #endregion

        #region what really matters 

        public void ExportToExcel()
        {
            foreach(KeyValuePair<string, ExcelWriter> kvp in Program.ExcelWriters) {
                kvp.Value.SaveExcelFile(kvp.Key);
            }
        }

        #endregion
    }
}
