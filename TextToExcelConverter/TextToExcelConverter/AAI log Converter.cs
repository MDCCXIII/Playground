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

            //mandatory. Otherwise will throw an exception when calling ReportProgress method  
            backgroundWorker1.WorkerReportsProgress = true;

            //mandatory. Otherwise we would get an InvalidOperationException when trying to cancel the operation  
            backgroundWorker1.WorkerSupportsCancellation = true;
        }

        InstanceConverter ic = null;
        DragEventArgs DEA = null;
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            DEA = e;
            backgroundWorker1.RunWorkerAsync();
        }

        protected virtual bool IsFileinUse(string file)
        {
            bool result = false;
            Stream s = null;
            try {
                s = File.Open(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException) {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                result = true;
            }
            finally {
                if (s != null)
                    s.Close();
            }
            return result;
        }
        
        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                e.Effect = DragDropEffects.Copy;
            } else {
                e.Effect = DragDropEffects.None;
            }
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            ConvertOptions.SelectedIndex = 0;
            setInfoLabelText("");

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start(@Directory.GetCurrentDirectory());
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(ic != null) {
                ic.Cleanup();
            }
            
        }



        public void setInfoLabelText(string text)
        {
            infoLabel.Text = text;
        }

        public void SetProgressBarValue(int value)
        {
            progressBar1.Value = value;
        }

        public int GetConversionOptionSelectedIndex()
        {
            return ConvertOptions.SelectedIndex;
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            if (DEA != null) {
                cancel.Visible = true;
                cancel.Enabled = true;
                progressBar1.Visible = true;
                ic = new InstanceConverter(DEA, this);
            }
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
            progressBar1.Visible = false;
            setInfoLabelText("Conversion Cancelled.");
            cancel.Visible = false;
            cancel.Enabled = false;
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            SetProgressBarValue(e.ProgressPercentage);
        }
    }
}
