using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AAI_Log_Converter.ExcelInterop;
using Excel = Microsoft.Office.Interop.Excel;
using AAI_Log_Converter.FileIO;
using System.IO;

namespace AAI_Log_Converter
{
    class FileConverter
    {
        private ControlSet controlSet;
        private Form1 form1 = null;
        internal string filePath;
        internal string FileName;
        private string newFileName;
        private int fileLineCount = 0;
        private List<string> file = new List<string>();
        private GenerateTimeLogWorksheet generateTimeLogWorksheet = new GenerateTimeLogWorksheet();
        private GenerateDataWorksheet generateDataWorksheet = new GenerateDataWorksheet();
        private GenerateUsageStatisticsWorksheet generateUsageStatisticsWorksheet = new GenerateUsageStatisticsWorksheet();

        public FileConverter(string filePath, Form1 form1, ControlSet controlSet)
        {
            try {
                this.form1 = form1;
                this.controlSet = controlSet;
                Init();
                if (!controlSet.CancellationPending()) {
                    if (FileUtils.FileExists(filePath)) {
                        this.filePath = filePath;
                        this.FileName = FileUtils.GetFileName(filePath);
                        if (!Program.FileNames.Contains(FileName)) {
                            Program.FileNames.Add(FileName);
                        }
                        newFileName = Directory.GetCurrentDirectory() + "\\" + FileName + ".xlsx";
                        if (!FileUtils.IsFileinUse(newFileName)) {
                            if (!controlSet.CancellationPending()) {
                                fileLineCount = FileUtils.GetFileLineCount(filePath);
                            }
                            if (!controlSet.CancellationPending()) {
                                file = FileUtils.FileToList(filePath);
                            }
                            if (!controlSet.CancellationPending()) {
                                ConvertFile();
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                //TODO: Add Logging
            }
        }
        
        private void ConvertFile()
        {
            if (file.Count > 0) {
                generateUsageStatisticsWorksheet = new GenerateUsageStatisticsWorksheet();
                int rowNumber = 1;
                int lineNumber = 0;
                foreach (string line in file) {
                    lineNumber++;
                    GenerateWorksheets(line.Trim(), rowNumber);
                    if (line.StartsWith("),")) {
                        rowNumber++;
                    }
                    controlSet.Report(((double)lineNumber / (double)fileLineCount) * 100);
                }
                generateUsageStatisticsWorksheet.OutputToExcel(this);
            }
        }

        private void GenerateWorksheets(string line, int rowNumber)
        {
            if (!controlSet.CancellationPending()) {
                if (controlSet.form1 != null) {
                    switch (controlSet.GetConversionOptionsSelectedIndex()) {
                        case 1:
                            GenerateData(rowNumber, line);
                            break;
                        case 2:
                            GenerateFieldUsageStatistics(rowNumber, line);
                            break;
                        default:
                            GenerateData(rowNumber, line);
                            GenerateFieldUsageStatistics(rowNumber, line);
                            break;
                    }
                } else {
                    GenerateData(rowNumber, line);
                    GenerateFieldUsageStatistics(rowNumber, line);
                }
            }
        }

        public void GenerateTimeLogs(int rowNumber, string line)
        {
            generateTimeLogWorksheet.GenerateTimeLog(rowNumber, line, this);
        }

        private void GenerateFieldUsageStatistics(int rowNumber, string line)
        {
            generateUsageStatisticsWorksheet.GenerateStatistics(rowNumber, line);
        }

        private void GenerateData(int rowNumber, string line)
        {
            generateDataWorksheet.formatText(rowNumber, line, this);
        }

        private void Init()
        {
            FileName = "";
            fileLineCount = 0;
            controlSet.SetProgressBarValue(0);
        }
        
    }
}
