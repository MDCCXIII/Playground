using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;

namespace TextToExcelConverter
{
    class ConvertingFile
    {
        const string beginConversionText = "Begin Conversion";
        const string pauseText = "Pause";
        const string resumeText = "Resume";


        Form1 form1;
        public Label FilePathLabel = new Label();
        public ProgressBar pb = new ProgressBar();
        public Button Cancel_Button = new Button();
        public Button StartPause_Button = new Button();
        public Label InfoLabel = new Label();
        BackgroundWorker bw = new BackgroundWorker();
        public ManualResetEvent _pauseEvent = new ManualResetEvent(true);
        InstanceConverter ic = null;
        public ComboBox convertOptions;

        public ConvertingFile(Form1 form1, string filePath)
        {
            this.form1 = form1;
            NewFilePathLabel(filePath);
            NewProgressBar();
            NewCancelButton();
            NewStartPause_Button();
            NewBackgroundWorker();
            NewConvertOptions();
        }

        private void NewConvertOptions()
        {
            convertOptions = new ComboBox();
            convertOptions.FormattingEnabled = true;
            convertOptions.Items.AddRange(new object[] {
            "All",
            "Data",
            "Usage Statistics"});
            convertOptions.Name = "convertOptions";
            convertOptions.Size = new System.Drawing.Size(112, 21);
            convertOptions.SelectedIndex = 0;
        }

        delegate void SetTextCallback(string text);
        private void ThreadSafeSetInfoLabelText(string text)
        {
            InfoLabel.Text = text;
        }
        public void setInfoLabelText(string text)
        {
            if (InfoLabel.InvokeRequired) {
                SetTextCallback d = new SetTextCallback(ThreadSafeSetInfoLabelText);
                form1.Invoke(d, new object[] { text });
            } else {
                ThreadSafeSetInfoLabelText(text);
            }

        }

        delegate void SetProgressCallback(int value);
        private void ThreadSafeSetProgressBarValue(int value)
        {
            pb.Value = value;
        }
        public void SetProgressBarValue(int value)
        {
            if (pb.InvokeRequired) {
                SetProgressCallback d = new SetProgressCallback(ThreadSafeSetProgressBarValue);
                form1.Invoke(d, new object[] { value });
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
                return (int)form1.Invoke(d);
            } else {
                return ThreadSafeGetConversionOptionSelectedIndex();
            }
        }

        delegate bool SetCancellationPendingCallback();
        private bool ThreadSafeGetCancellationPending()
        {
            return bw.CancellationPending;
        }
        public bool GetCancellationPending()
        {
            SetCancellationPendingCallback d = new SetCancellationPendingCallback(ThreadSafeGetCancellationPending);
            return (bool)form1.Invoke(d);
        }

        private void NewBackgroundWorker()
        {
            bw = new BackgroundWorker();
            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;
            bw.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            bw.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ic = new InstanceConverter(FilePathLabel.Text, form1, this, e);
            if (bw.CancellationPending) {
                e.Cancel = true;
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Debug.WriteLine(e);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled) {
                setInfoLabelText("Conversion Cancelled.");
                SetProgressBarValue(100);
            } else if (e.Error != null) {
                setInfoLabelText("There was an error during the conversion.");
            } else {
                setInfoLabelText("Conversion Complete.");
            }

            StartPause_Button.Text = beginConversionText;
        }

        private void NewInfoLabelLabel(string filePath)
        {
            InfoLabel = new Label();
            setInfoLabelText("");
            InfoLabel.Width = 350;
            InfoLabel.Anchor = AnchorStyles.Left;
            InfoLabel.AutoSize = false;
            InfoLabel.AutoEllipsis = true;
        }

        private void NewStartPause_Button()
        {
            StartPause_Button = new Button();
            StartPause_Button.Text = beginConversionText;
            StartPause_Button.AutoSize = true;
            StartPause_Button.UseVisualStyleBackColor = true;
            StartPause_Button.Click += new EventHandler(startConversion_Click);
            StartPause_Button.Anchor = AnchorStyles.Left;
        }

        private void NewCancelButton()
        {
            Cancel_Button = new Button();
            Cancel_Button.Text = "Cancel";
            Cancel_Button.AutoSize = true;
            Cancel_Button.UseVisualStyleBackColor = true;
            Cancel_Button.Click += new EventHandler(cancel_Click);
            Cancel_Button.Anchor = AnchorStyles.Left;
        }

        private void NewFilePathLabel(string filePath)
        {
            FilePathLabel = new Label();
            FilePathLabel.Text = filePath;
            FilePathLabel.AutoSize = true;
            FilePathLabel.Anchor = AnchorStyles.Left;
            FilePathLabel.AutoSize = false;
            FilePathLabel.AutoEllipsis = true;
        }

        private void NewProgressBar()
        {
            pb = new ProgressBar();
            pb.Size = new System.Drawing.Size(200, 23);
            pb.Anchor = AnchorStyles.Left;
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            if (bw.IsBusy) {
                bw.CancelAsync();
                if (bw.CancellationPending) {
                    setInfoLabelText("Cancelling conversion...");
                }
            }
        }

        private void startConversion_Click(object sender, EventArgs e)
        {
            if (StartPause_Button.Text.Equals(beginConversionText)) {
                StartProccess();
            } else if (StartPause_Button.Text.Equals(pauseText)) {
                Pause();
                StartPause_Button.Text = resumeText;
            } else if (StartPause_Button.Text.Equals(resumeText)) {
                Resume();
                StartPause_Button.Text = pauseText;
            }
        }

        public void StartProccess()
        {
            bw.RunWorkerAsync();
            StartPause_Button.Text = pauseText;
        }

        private void Pause()
        {
            _pauseEvent.Reset();
        }

        private void Resume()
        {
            _pauseEvent.Set();
        }

        public void wait()
        {
            _pauseEvent.WaitOne(Timeout.Infinite);
        }
            
    }
}
