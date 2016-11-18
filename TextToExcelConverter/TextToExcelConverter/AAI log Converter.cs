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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        static Excel.Application xlApp;
        public static Excel.Workbook wb;
        public static Excel.Worksheet ws;
        public static Excel.Range rng;
        static string filePath;
        public static string fileName;
        public static Dictionary<string, int> columns = new Dictionary<string, int>();
        static int totalLines = 0;
        List<string> file = new List<string>();

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            try {
                Init();
                if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                    string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
                    foreach (string filePath in filePaths) {
                        if (Exists(filePath)) {
                            CreateNewWorkBook();
                            SetFilePath(filePath);
                            if (IsTextFile(filePath)) {
                                SetFileName(filePath);
                                SetTotalLines();
                                ParseFile();
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

        protected virtual bool IsFileinUse(FileInfo file)
        {
            bool result = false;
            FileStream stream = null;
            try {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException) {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                result = true;
            }
            finally {
                if (stream != null)
                    stream.Close();
            }
            return result;
        }

        private static bool Exists(string filePath)
        {
            return File.Exists(filePath);
        }

        private static bool IsTextFile(string filePath)
        {
            return filePath.EndsWith(".txt");
        }

        private static void SetFilePath(string filePath)
        {
            Form1.filePath = filePath;
        }

        private static void SetFileName(string fileLoc)
        {
            fileName = Path.GetFileName(fileLoc).Replace(".txt", "");
        }

        private static void SetTotalLines()
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
                if (comboBox1.SelectedIndex.Equals(0) || comboBox1.SelectedIndex.Equals(1)) {
                    progressBar1.Value = 0;
                    CreateOrActivatWorkSheet("Data");
                    TextToExcel();
                }
                if (comboBox1.SelectedIndex.Equals(0) || comboBox1.SelectedIndex.Equals(2)) {
                    progressBar1.Value = 0;
                    CreateOrActivatWorkSheet("Usage Statistics");
                    FieldUsage();
                    
                }
                progressBar1.Value = 0;
                CreateOrActivatWorkSheet("Service Call Time Log");
                TimeLogs();
                SaveExcelFile();
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
            TextFileFormatter.lastRowColumnFoundIn.Clear();
            TextFileFormatter.timesSeenInRow.Clear();
            Statistics.lastRowColumnFoundIn.Clear();
            Statistics.timesSeenInRow.Clear();
            Statistics.fieldSeenTotal.Clear();
            Statistics.fieldNullCount.Clear();
            progressBar1.Value = 0;
        }

        private static void CreateNewWorkBook()
        {
            wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
        }

        private static void CreateOrActivatWorkSheet(string wsName)
        {
            if (ExistOrCreate(wsName)) {
                if (Form1.ws.Name == wsName) {
                    Form1.wb.Worksheets[wsName].Activate();
                    Form1.ws = Form1.wb.ActiveSheet;
                }
            }

        }

        private static bool ExistOrCreate(string wsName)
        {
            try {
                Excel.Worksheet t = Form1.wb.Worksheets[wsName];
            }
            catch (System.Runtime.InteropServices.COMException) {
                try {
                    var xlSheets = Form1.wb.Sheets as Excel.Sheets;
                    var xlNewSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
                    xlNewSheet.Name = wsName;
                    Form1.ws = xlNewSheet;

                }
                catch (Exception e) {
                    Debug.WriteLine(e);
                }

            }
            return true;
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
                catch (System.Runtime.InteropServices.COMException e) {
                    failed = true;
                }
                System.Threading.Thread.Sleep(10);
            } while (failed && attempt > 3);
            if (failed) {
                progressBar1.Value = 0;
                label2.Text = "Failed to save the converted file.";
            } else {
                progressBar1.Value = 100;
                label2.Text = "Conversion Complete";
            }
            Cleanup();
            
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                e.Effect = DragDropEffects.Copy;
            } else {
                e.Effect = DragDropEffects.None;
            }
        }

        private void TextToExcel()
        {
            int rowNumber = 1;
            int lineNumber = 0;
            string newLine;
            label2.Text = "Converting Data for File: " + fileName + ".txt";
            foreach (string line in file) {
                lineNumber++;
                double progress = ((double)lineNumber / (double)totalLines) * 100;
                Report(progress);
                newLine = line.Trim();
                TextFileFormatter.formatText(ref rowNumber, line);
            }

        }

        private void FieldUsage()
        {
            int rowNumber = 1;
            int lineNumber = 0;
            string newLine;
            string previousLine = null;
            label2.Text = "Generating Usage Statistics for File: " + fileName + ".txt";
            foreach (string line in file) {
                if (previousLine != null) {
                    lineNumber++;
                    double progress = ((double)lineNumber / (double)totalLines) * 100;
                    Report(progress);
                    newLine = line.Trim();
                    Statistics.GenerateStatistics(ref rowNumber, line);
                }
                previousLine = line;
            }
            Statistics.OutputToExcel();
        }

        private void TimeLogs()
        {
            int rowNumber = 1;
            int lineNumber = 0;
            string newLine;
            string previousLine = null;
            label2.Text = "Generating Time Logs for File: " + fileName + ".txt";
            foreach (string line in file) {
                if (previousLine != null) {
                    lineNumber++;
                    double progress = ((double)lineNumber / (double)totalLines) * 100;
                    Report(progress);
                    newLine = line.Trim();
                    TimeLog.GenerateTimeLog(ref rowNumber, line);
                }
                previousLine = line;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 0;
            label2.Text = "";

        }

        public void Report(double progress)
        {
            int posistion = (int)Math.Round(progress);
            if(posistion <= 100) {
                progressBar1.Value = posistion;
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start(@Directory.GetCurrentDirectory());
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Cleanup();
        }

        private static void Cleanup()
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
