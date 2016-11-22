using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    class InstanceConverter
    {
        Excel.Application xlApp;
        public Excel.Workbook wb;
        public Excel.Worksheet ws;
        public Excel.Range rng;
        string filePath;
        public string fileName;
        public Dictionary<string, int> columns = new Dictionary<string, int>();
        int totalLines = 0;
        List<string> file = new List<string>();
        Form1 Form1;
        ConvertingFile cf;
        bool cancel = false;

        public InstanceConverter(string file, Form1 Form1, ConvertingFile cf, System.ComponentModel.DoWorkEventArgs dwea)
        {
            try {
                this.Form1 = Form1;
                this.cf = cf;
                Init();
                if (!CancellationPending()) {
                    if (Exists(file)) {
                        CreateNewWorkBook();
                        SetFilePath(file);
                        SetFileName(file);
                        if (!IsFileinUse(Directory.GetCurrentDirectory() + "\\" + fileName + ".xlsx")) {
                            if (!CancellationPending()) {
                                SetTotalLines();
                            }
                            if (!CancellationPending()) {
                                ParseFile();
                            }
                            if (!CancellationPending()) {
                                ConvertFile();
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                Debug.WriteLine(ex.Message);
            }
        }

        private bool CancellationPending()
        {
            cf._pauseEvent.WaitOne(Timeout.Infinite);
            if (!cancel) {
                cancel = cf.GetCancellationPending();
            }
            return cancel;
        }

        private void Init()
        {
            xlApp = new Excel.Application();
            if (xlApp == null) {
                MessageBox.Show("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            wb = null;
            ws = null;
            rng = null;
            fileName = "";
            columns.Clear();
            totalLines = 0;
            cf.SetProgressBarValue(0);
        }

        private static bool Exists(string filePath)
        {
            return File.Exists(filePath);
        }

        private void CreateNewWorkBook()
        {
            wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
        }

        private void SetFilePath(string filePath)
        {
            this.filePath = filePath;
        }

        protected virtual bool IsFileinUse(string file)
        {
            bool result = false;
            Stream s = null;
            try {
                if (Exists(file)) {
                    s = File.Open(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
            }
            catch (IOException) {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                result = true;
                MessageBox.Show("The File " + file + " is currently open.\nPlease close the file and try again.");
            }
            finally {
                if (s != null) {
                    s.Close();
                    File.Delete(file);
                }
            }
            return result;
        }

        private void SetFileName(string fileLoc)
        {
            fileName = Path.GetFileName(fileLoc).Replace(".txt", "");
        }

        private void SetTotalLines()
        {
            totalLines = File.ReadAllLines(filePath).Count();
        }

        private void ParseFile()
        {
            using (TextReader tr = new StreamReader(filePath)) {
                string line;
                while ((line = tr.ReadLine()) != null) {
                    file.Add(line);
                }
                tr.Close();
            }
        }

        private void ConvertFile()
        {
            if (file.Count > 0) {
                if (!CancellationPending()) {
                    if (cf.GetConversionOptionSelectedIndex() == 0 || cf.GetConversionOptionSelectedIndex() == 1) {
                        cf.SetProgressBarValue(0);
                        CreateOrActivatWorkSheet("Data");
                        TextToExcel();
                    }
                }
                if (!CancellationPending()) {
                    if (cf.GetConversionOptionSelectedIndex() == 0 || cf.GetConversionOptionSelectedIndex() == 2) {
                        cf.SetProgressBarValue(0);
                        CreateOrActivatWorkSheet("Usage Statistics");
                        FieldUsage();

                    }
                }
                if (!CancellationPending()) {
                    cf.SetProgressBarValue(0);
                    CreateOrActivatWorkSheet("Service Call Time Log");
                    TimeLogs();
                }
                if (!CancellationPending()) {
                    SaveExcelFile();
                }
            }
        }

        private void CreateOrActivatWorkSheet(string wsName)
        {
            if (ExistOrCreate(wsName)) {
                if (ws.Name == wsName) {
                    wb.Worksheets[wsName].Activate();
                    ws = wb.ActiveSheet;
                }
            }

        }

        private void TextToExcel()
        {
            TextFileFormatter tff = new TextFileFormatter();
            int rowNumber = 1;
            int lineNumber = 0;
            string newLine;
            cf.setInfoLabelText("Converting Data for File: " + fileName + ".txt");
            foreach (string line in file) {
                if (!CancellationPending()) {
                    lineNumber++;
                    double progress = ((double)lineNumber / (double)totalLines) * 100;
                    Report(progress);
                    newLine = line.Trim();
                    tff.formatText(ref rowNumber, line, this);
                }
            }

        }

        private void FieldUsage()
        {
            Statistics stats = new Statistics();
            int rowNumber = 1;
            int lineNumber = 0;
            string newLine;
            string previousLine = null;
            cf.setInfoLabelText("Generating Usage Statistics for File: " + fileName + ".txt");
            foreach (string line in file) {
                if (!CancellationPending()) {
                    if (previousLine != null) {
                        lineNumber++;
                        double progress = ((double)lineNumber / (double)totalLines) * 100;
                        Report(progress);
                        newLine = line.Trim();
                        stats.GenerateStatistics(ref rowNumber, line, this);
                    }
                    previousLine = line;
                }
            }
            stats.OutputToExcel(this);
        }

        private void TimeLogs()
        {
            TimeLog tl = new TimeLog();
            int rowNumber = 1;
            int lineNumber = 0;
            string newLine;
            string previousLine = null;
            cf.setInfoLabelText("Generating Time Logs for File: " + fileName + ".txt");
            rng = (Excel.Range)wb.ActiveSheet.Cells[1, 1];
            rng.Value = "Service Call #";
            rng = (Excel.Range)wb.ActiveSheet.Cells[1, 2];
            rng.Value = "Date Time";
            foreach (string line in file) {
                if (!CancellationPending()) {
                    if (previousLine != null) {
                        lineNumber++;
                        double progress = ((double)lineNumber / (double)totalLines) * 100;
                        Report(progress);
                        newLine = line.Trim();
                        tl.GenerateTimeLog(ref rowNumber, line, this);
                    }
                    previousLine = line;
                }
            }
        }

        private void SaveExcelFile()
        {
            int attempt = 0;
            bool failed = false;
            if (!CancellationPending()) {
                do {
                    if (!CancellationPending()) {
                        try {
                            attempt++;
                            wb.SaveAs(Directory.GetCurrentDirectory() + "\\" + fileName + ".xlsx");
                            failed = false;
                        }
                        catch (System.Runtime.InteropServices.COMException) {
                            failed = true;
                        }
                        System.Threading.Thread.Sleep(10);
                    }
                } while (failed && attempt > 3);
                if (failed) {
                    cf.SetProgressBarValue(0);
                    cf.setInfoLabelText("Failed to save the converted file.");
                } else {
                    cf.SetProgressBarValue(100);
                    cf.setInfoLabelText("Conversion Complete.");
                }
            }
            if (!CancellationPending()) {
                Cleanup(true);
            } else {
                Cleanup(false);
            }
        }

        private bool ExistOrCreate(string wsName)
        {
            try {
                Excel.Worksheet t = wb.Worksheets[wsName];
            }
            catch (System.Runtime.InteropServices.COMException) {
                try {
                    var xlSheets = wb.Sheets as Excel.Sheets;
                    var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                    xlNewSheet.Name = wsName;
                    ws = xlNewSheet;

                }
                catch (Exception e) {
                    Debug.WriteLine(e);
                }

            }
            return true;
        }

        public void Report(double progress)
        {
            int posistion = (int)Math.Round(progress);
            if (posistion <= 100) {
                cf.SetProgressBarValue(posistion);
            }

        }

        public void Cleanup(bool save)
        {
            try {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(rng);
                Marshal.FinalReleaseComObject(ws);

                wb.Close(save, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(wb);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
            }
            catch (Exception) {

            }
        }
    }
}
