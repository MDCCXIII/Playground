using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace TextToExcelConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        
        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
                foreach (string fileLoc in filePaths) {
                    // Code to read the contents of the text file
                    if (File.Exists(fileLoc)) {
                        CreateExcelFile(fileLoc);
                        using (TextReader tr = new StreamReader(fileLoc)) {
                            TextToExcel(tr);

                        }
                    }

                }
            }
        }

        private void CreateExcelFile(string fileLoc)
        {
            Path.GetFileName(fileLoc);
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                e.Effect = DragDropEffects.Copy;
            } else {
                e.Effect = DragDropEffects.None;
            }
        }

        private static void TextToExcel(TextReader tr)
        {
            MessageBox.Show(tr.ReadToEnd());
            string line;
            while ((line = tr.ReadLine()) != null) {
               
            }

        }
    }
}
