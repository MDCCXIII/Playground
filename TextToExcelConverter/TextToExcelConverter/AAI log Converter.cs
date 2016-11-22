using System;
using System.Collections.Generic;
using System.ComponentModel;
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
        }

        InstanceConverter ic = null;
        List<string> FilePath = new List<string>();
        int rowNumber = 0;

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
                foreach (string filePath in filePaths) {
                    if (filePath.EndsWith(".txt")) {
                        if (!FilePath.Contains(filePath)) {
                            FilePath.Add(filePath);
                            ConvertingFile cf = new ConvertingFile(this, filePath);
                            AddNewRow(cf);
                            cf.StartProccess();
                        }
                    }
                }
            }
        }

        private void AddNewRow(ConvertingFile cf)
        {
            tableLayoutPanel1.RowCount = rowNumber + 1;
            tableLayoutPanel1.Controls.Add(cf.FilePathLabel, 0, rowNumber);
            tableLayoutPanel1.Controls.Add(cf.convertOptions, 1, rowNumber);
            tableLayoutPanel1.Controls.Add(cf.StartPause_Button, 2, rowNumber);
            tableLayoutPanel1.Controls.Add(cf.InfoLabel, 3, rowNumber);
            tableLayoutPanel1.Controls.Add(cf.pb, 4, rowNumber);
            tableLayoutPanel1.Controls.Add(cf.Cancel_Button, 5, rowNumber);
            rowNumber++;
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
            RemoveEmptyRows();
        }

        private void RemoveEmptyRows()
        {
            for (int row = tableLayoutPanel1.RowCount - 1; row >= 0; row--) {
                bool hasControl = false;
                for (int col = 0; col < tableLayoutPanel1.ColumnCount; col++) {
                    if (tableLayoutPanel1.GetControlFromPosition(col, row) != null) {
                        hasControl = true;
                        break;
                    }
                }

                if (!hasControl) {
                    tableLayoutPanel1.RowStyles.RemoveAt(row);
                    tableLayoutPanel1.RowCount--;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start(@Directory.GetCurrentDirectory());
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (ic != null) {
                ic.Cleanup(false);
            }

        }

    }
}
