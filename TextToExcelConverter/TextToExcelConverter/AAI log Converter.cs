using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

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
            convertOptions.SelectedIndex = 0;
            setInfoLabelText("");

        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start(@Directory.GetCurrentDirectory());
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(ic != null) {
                ic.Cleanup(false);
            }
            
        }

        delegate void SetTextCallback(string text);
        private void ThreadSafeSetInfoLabelText(string text)
        {
            this.infoLabel.Text = text;
        }
        public void setInfoLabelText(string text)
        {
            if (infoLabel.InvokeRequired) {
                SetTextCallback d = new SetTextCallback(ThreadSafeSetInfoLabelText);
                this.Invoke(d, new object[] { text });
            } else {
                ThreadSafeSetInfoLabelText(text);
            }
            
        }

        delegate void SetProgressCallback(int value);
        private void ThreadSafeSetProgressBarValue(int value)
        {
            progressBar1.Value = value;
        }
        public void SetProgressBarValue(int value)
        {
            if (progressBar1.InvokeRequired) {
                SetProgressCallback d = new SetProgressCallback(ThreadSafeSetProgressBarValue);
                this.Invoke(d, new object[] { value });
            } else {
                ThreadSafeSetProgressBarValue(value);
            }
        }

        delegate int SetOptionSelectedCallback();
        private int ThreadSafeGetConversionOptionSelectedIndex()
        {
            return convertOptions.SelectedIndex;
        }
        public int GetConversionOptionSelectedIndex()
        {
            if (convertOptions.InvokeRequired) {
                SetOptionSelectedCallback d = new SetOptionSelectedCallback(ThreadSafeGetConversionOptionSelectedIndex);
                return (int)this.Invoke(d);
            } else {
                return ThreadSafeGetConversionOptionSelectedIndex();
            }
        }

        delegate bool SetCancellationPendingCallback();
        private bool ThreadSafeGetCancellationPending()
        {
            return backgroundWorker1.CancellationPending;
        }
        public bool GetCancellationPending()
        {
            SetCancellationPendingCallback d = new SetCancellationPendingCallback(ThreadSafeGetCancellationPending);
                return (bool)this.Invoke(d);
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            if (DEA != null) {
                ic = new InstanceConverter(DEA, this, e);
                if (backgroundWorker1.CancellationPending) {
                    e.Cancel = true;
                }
            }
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
            if (backgroundWorker1.CancellationPending) {
                setInfoLabelText("Cancelling conversion...");
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
        {
            //SetProgressBarValue(e.ProgressPercentage);
            Debug.WriteLine(e);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled) {
                setInfoLabelText("Conversion Cancelled.");
                SetProgressBarValue(100);
            } else if (e.Error != null) {
                setInfoLabelText("There was an error during the conversion.");
            } else {
                setInfoLabelText("Conversion Complete.");
            }
        }
    }
}
