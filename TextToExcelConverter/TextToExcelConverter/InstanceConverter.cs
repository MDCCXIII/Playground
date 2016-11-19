using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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

        public InstanceConverter(DragEventArgs e, Form1 Form1)
        {
            try {
                this.Form1 = Form1;
                Init();
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                    string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
                    foreach (string filePath in filePaths) {
                        if (Exists(filePath)) {
                            CreateNewWorkBook();
                            SetFilePath(filePath);
                            if (IsTextFile(filePath)) {
                                SetFileName(filePath);
                                if(!IsFileinUse(Directory.GetCurrentDirectory() + "\\" + fileName + ".xlsx")) {
                                    SetTotalLines();
                                    ParseFile();
                                    ConvertFile();
                                }
                               
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                Debug.WriteLine(ex.Message);
            }
        }

        private void Init()
        {
            xlApp = new Excel.Application();
            if (xlApp == null) {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            wb = null;
            ws = null;
            rng = null;
            filePath = "";
            fileName = "";
            columns.Clear();
            totalLines = 0;
            //Form1.SetProgressBarValue(0);
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

        private bool IsTextFile(string filePath)
        {
            return filePath.EndsWith(".txt");
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
                if (Form1.GetConversionOptionSelectedIndex() == 0 || Form1.GetConversionOptionSelectedIndex() == 1) {
                    //Form1.SetProgressBarValue(0);
                    CreateOrActivatWorkSheet("Data");
                    TextToExcel();
                }
                if (Form1.GetConversionOptionSelectedIndex() == 0 || Form1.GetConversionOptionSelectedIndex() == 2) {
                    //Form1.SetProgressBarValue(0);
                    CreateOrActivatWorkSheet("Usage Statistics");
                    FieldUsage();

                }
                //Form1.SetProgressBarValue(0);
                CreateOrActivatWorkSheet("Service Call Time Log");
                TimeLogs();
                SaveExcelFile();
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
            Form1.setInfoLabelText("Converting Data for File: " + fileName + ".txt");
            foreach (string line in file) {
                lineNumber++;
                double progress = ((double)lineNumber / (double)totalLines) * 100;
                //Report(progress);
                newLine = line.Trim();
                tff.formatText(ref rowNumber, line, this);
            }

        }

        private void FieldUsage()
        {
            Statistics stats = new Statistics();
            int rowNumber = 1;
            int lineNumber = 0;
            string newLine;
            string previousLine = null;
            Form1.setInfoLabelText("Generating Usage Statistics for File: " + fileName + ".txt");
            foreach (string line in file) {
                if (previousLine != null) {
                    lineNumber++;
                    double progress = ((double)lineNumber / (double)totalLines) * 100;
                    //Report(progress);
                    newLine = line.Trim();
                    stats.GenerateStatistics(ref rowNumber, line, this);
                }
                previousLine = line;
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
            Form1.setInfoLabelText("Generating Time Logs for File: " + fileName + ".txt");
            foreach (string line in file) {
                if (previousLine != null) {
                    lineNumber++;
                    double progress = ((double)lineNumber / (double)totalLines) * 100;
                    //Report(progress);
                    newLine = line.Trim();
                    tl.GenerateTimeLog(ref rowNumber, line, this);
                }
                previousLine = line;
            }
        }

        private void SaveExcelFile()
        {
            int attempt = 0;
            bool failed = false;
            do {
                try {
                    attempt++;
                    wb.SaveAs(Directory.GetCurrentDirectory() + "\\" + fileName + ".xlsx");
                    failed = false;
                }
                catch (System.Runtime.InteropServices.COMException) {
                    failed = true;
                }
                System.Threading.Thread.Sleep(10);
            } while (failed && attempt > 3);
            if (failed) {
                //Form1.SetProgressBarValue(0);
                Form1.setInfoLabelText("Failed to save the converted file.");
            } else {
                //Form1.SetProgressBarValue(100);
                Form1.setInfoLabelText("Conversion Complete");
            }
            Cleanup();

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
                Form1.SetProgressBarValue(posistion);
            }

        }

        public void Cleanup()
        {
            try {
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.FinalReleaseComObject(rng);
                Marshal.FinalReleaseComObject(ws);

                wb.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(wb);

                xlApp.Quit();
                Marshal.FinalReleaseComObject(xlApp);
            }
            catch (Exception) {

            }
        }
    }
}
