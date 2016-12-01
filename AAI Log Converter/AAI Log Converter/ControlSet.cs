using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AAI_Log_Converter.ExcelInterop
{
    class ControlSet : IDisposable
    {
        #region Globals

        private const string START_TEXT = "Begin Conversion";
        private const string PAUSE_TEXT = "Pause";
        private const string RESUME_TEXT = "Resume";

        internal Form1 form1 = null;
        private BackgroundWorker backgroundWorker = new BackgroundWorker();
        private FileConverter fileConverter = null;

        private bool cancel = false;
        private string filePath;

        internal Label filePath_Label = null;
        internal ProgressBar progressBar = null;
        internal Button cancel_Button = null;
        internal Button start_Button = null;
        internal Label info_Label = null;
        internal ComboBox convert_ComboBox = null;
        internal ManualResetEvent _pauseEvent = new ManualResetEvent(true);

        #region Delegates

        private delegate void SetTextCallback(string text);
        private delegate void SetProgressCallback(int value);
        private delegate int SetOptionSelectedCallback();
        private delegate bool SetCancellationPendingCallback();

        #endregion

        #endregion

        #region Constructors

        //Allow new threads to start without gui interface (e.g. command line)
        public ControlSet(string filePath)
        {
            this.filePath = filePath;
            NewBackgroundWorker();
        }

        //constructor for windows form gui
        public ControlSet(Form1 form1, string filePath)
        {
            this.form1 = form1;
            this.filePath = filePath;
            NewFilePathLabel(filePath);
            NewInfoLabelLabel();
            NewProgressBar();
            NewCancelButton();
            NewStartPause_Button();
            NewBackgroundWorker();
            NewConvertOptions();
        }

        #endregion

        #region Accessors
        
        public int GetConversionOptionsSelectedIndex()
        {
            if(form1 != null) {
                if (convert_ComboBox.InvokeRequired) {
                    SetOptionSelectedCallback d = new SetOptionSelectedCallback(ThreadSafeGetConversionOptionSelectedIndex);
                    return (int)form1.Invoke(d);
                } else {
                    return ThreadSafeGetConversionOptionSelectedIndex();
                }
            }
            return 0;
        }

        public void SetProgressBarValue(int value)
        {
            if (form1 != null) {
                if (progressBar.InvokeRequired) {
                    SetProgressCallback d = new SetProgressCallback(ThreadSafeSetProgressBarValue);
                    form1.Invoke(d, new object[] { value });
                } else {
                    ThreadSafeSetProgressBarValue(value);
                }
            }
        }

        public void setInfoLabelText(string text)
        {
            if (form1 != null) {
                if (info_Label.InvokeRequired) {
                    SetTextCallback d = new SetTextCallback(ThreadSafeSetInfoLabelText);
                    form1.Invoke(d, new object[] { text });
                } else {
                    ThreadSafeSetInfoLabelText(text);
                }
            }
        }

        public void StartProccess()
        {
            backgroundWorker.RunWorkerAsync();
            if (form1 != null) {
                start_Button.Text = PAUSE_TEXT;
            }
        }

        public void Report(double progress)
        {
            int posistion = (int)Math.Round(progress);
            if (posistion <= 100) {
                SetProgressBarValue(posistion);
            }

        }

        public bool CancellationPending()
        {
            _pauseEvent.WaitOne(Timeout.Infinite);
            if (!cancel) {
                cancel = GetCancellationPending();
            }
            return cancel;
        }

        public void Dispose()
        {
            ((IDisposable)backgroundWorker).Dispose();
            ((IDisposable)cancel_Button).Dispose();
            ((IDisposable)start_Button).Dispose();
            ((IDisposable)filePath_Label).Dispose();
            ((IDisposable)progressBar).Dispose();
            ((IDisposable)info_Label).Dispose();
            ((IDisposable)convert_ComboBox).Dispose();
            _pauseEvent.Dispose();
        }

        #endregion

        #region Controls

        private void NewConvertOptions()
        {
            convert_ComboBox = new ComboBox();
            convert_ComboBox.FormattingEnabled = true;
            convert_ComboBox.Items.AddRange(new object[] {
            "All",
            "Data",
            "Usage Statistics"});
            convert_ComboBox.Name = "convertOptions";
            convert_ComboBox.Size = new System.Drawing.Size(112, 21);
            convert_ComboBox.SelectedIndex = 0;
            convert_ComboBox.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        }
        
        private void NewBackgroundWorker()
        {
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);
        }
        
        private void NewInfoLabelLabel()
        {
            info_Label = new Label();
            info_Label.Name = "InfoLabel";
            setInfoLabelText("");
            info_Label.Width = 350;
            info_Label.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
            info_Label.AutoSize = true;
            info_Label.AutoEllipsis = true;
        }

        private void NewStartPause_Button()
        {
            start_Button = new Button();
            start_Button.Name = "StartButton";
            start_Button.Text = START_TEXT;
            start_Button.AutoSize = true;
            start_Button.UseVisualStyleBackColor = true;
            start_Button.Click += new EventHandler(startConversion_Click);
            start_Button.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        }

        private void NewCancelButton()
        {
            cancel_Button = new Button();
            cancel_Button.Name = "CancelButton";
            cancel_Button.Text = "Cancel";
            cancel_Button.AutoSize = true;
            cancel_Button.UseVisualStyleBackColor = true;
            cancel_Button.Click += new EventHandler(cancel_Click);
            cancel_Button.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        }

        private void NewFilePathLabel(string filePath)
        {
            filePath_Label = new Label();
            filePath_Label.Name = "FilePathLabel";
            filePath_Label.Text = filePath;
            filePath_Label.Width = 350;
            filePath_Label.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            filePath_Label.AutoSize = true;
            filePath_Label.AutoEllipsis = true;
            filePath_Label.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        }

        private void NewProgressBar()
        {
            progressBar = new ProgressBar();
            progressBar.Name = "ProgressBar";
            progressBar.Size = new System.Drawing.Size(200, 23);
            progressBar.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        }

        #endregion

        #region Internals

        private void Pause()
        {
            _pauseEvent.Reset();
        }

        private void Resume()
        {
            _pauseEvent.Set();
        }

        private bool GetCancellationPending()
        {
            SetCancellationPendingCallback d = new SetCancellationPendingCallback(ThreadSafeGetCancellationPending);
            return (bool)form1.Invoke(d);
        }
        
        private void ThreadSafeSetInfoLabelText(string text)
        {
            if(info_Label != null) {
                info_Label.Text = text;
            }
        }
        
        
        private void ThreadSafeSetProgressBarValue(int value)
        {
            if (progressBar != null) {
                progressBar.Value = value;
            }
        }
        
        
        private int ThreadSafeGetConversionOptionSelectedIndex()
        {
            int result = 0;
            if(convert_ComboBox != null) {
                result = convert_ComboBox.SelectedIndex;
            }
            return result;
        }
        
       
        private bool ThreadSafeGetCancellationPending()
        {
            return backgroundWorker.CancellationPending;
        }

        #endregion

        #region Events

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            if (form1 != null) {
                fileConverter = new FileConverter(filePath, form1, this);
                if (backgroundWorker.CancellationPending) {
                    e.Cancel = true;
                }
            } else {
                fileConverter = new FileConverter(filePath, null, this);
            }
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (form1 != null) {
                if (e.Cancelled) {
                    setInfoLabelText("Conversion Cancelled.");
                    SetProgressBarValue(100);
                    SetProgressBarValue(0);
                    start_Button.Text = START_TEXT;
                } else if (e.Error != null) {
                    setInfoLabelText("There was an error during the conversion.");
                    SetProgressBarValue(0);
                    start_Button.Text = START_TEXT;
                } else {
                    setInfoLabelText("Conversion Complete.");
                    start_Button.Visible = false;
                }
            }

            
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            if (backgroundWorker.IsBusy) {
                backgroundWorker.CancelAsync();
                if (backgroundWorker.CancellationPending) {
                    setInfoLabelText("Cancelling conversion...");
                }
            }
        }

        private void startConversion_Click(object sender, EventArgs e)
        {
            if (start_Button.Text.Equals(START_TEXT)) {
                StartProccess();
            } else if (start_Button.Text.Equals(PAUSE_TEXT)) {
                Pause();
                start_Button.Text = RESUME_TEXT;
            } else if (start_Button.Text.Equals(RESUME_TEXT)) {
                Resume();
                start_Button.Text = PAUSE_TEXT;
            }
        }

        #endregion
    }
}
