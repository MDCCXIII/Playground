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

namespace Playground.Text_to_excel_converter
{
    public partial class TxtToExcel_Form : Form
    {
        public TxtToExcel_Form()
        {
            InitializeComponent();
        }

        private void TxtToExcel_Form_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
                //object data = e.Data.GetData(DataFormats.FileDrop);
                foreach (string fileLoc in filePaths) {
                    // Code to read the contents of the text file
                    if (File.Exists(fileLoc)) {
                        using (TextReader tr = new StreamReader(fileLoc)) {
                            MessageBox.Show(tr.ReadToEnd());
                        }
                    }

                }
            }
        }

        private void TxtToExcel_Form_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                e.Effect = DragDropEffects.Copy;
            } 
            //else {
            //    e.Effect = DragDropEffects.None;
            //}
        }

        private void TxtToExcel_Form_Load(object sender, EventArgs e)
        {
            this.AllowDrop = true;
        }
    }
}
