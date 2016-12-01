using AAI_Log_Converter.ExcelInterop;
using AAI_Log_Converter.FileIO;
using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace AAI_Log_Converter
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        private List<ControlSet> ControlSets = new List<ControlSet>();
        private Dictionary<string, List<string>> services = new Dictionary<string, List<string>>();
        private Dictionary<string, TableLayoutPanel> serviceTables = new Dictionary<string, TableLayoutPanel>();

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) {
                services.Clear();
                serviceTables.Clear();
                AddFilesToServiceCollection(e);
                ServiceCollectionToForm();
            }
        }
        
        private void ServiceCollectionToForm()
        {
            foreach (string key in services.Keys) {
                AddTab(key);
                BuildServiceTable(key);
                PopulateControlSetsToServiceTable(key);
                AddServiceTablesToTab(key);
            }
        }

        private void AddServiceTablesToTab(string key)
        {
            Panel p = NewPanel();
            if (tabControl1.TabPages[key].Controls.ContainsKey("LayoutPanel")) {
                TableLayoutPanel tableLayoutPanel = (TableLayoutPanel)tabControl1.TabPages[key].Controls["LayoutPanel"].Controls["controlSet"];
                tableLayoutPanel.RowCount = tableLayoutPanel.RowCount + serviceTables[key].RowCount;
                int count = serviceTables[key].Controls.Count;
                for (int i = 0; i < count; i++) {
                    tableLayoutPanel.Controls.Add(serviceTables[key].Controls[0], serviceTables[key].GetColumn(serviceTables[key].Controls[0]), tableLayoutPanel.RowCount - (serviceTables[key].RowCount - 1 - serviceTables[key].GetRow(serviceTables[key].Controls[0])));
                }
            } else {
                p.Controls.Add(serviceTables[key]);
                tabControl1.TabPages[key].Controls.Add(p);
            }
            tabControl1.TabPages[key].Controls["LayoutPanel"].Controls["controlSet"].Show();
        }

        private Panel NewPanel()
        {
            Panel p = new Panel();
            p.Name = "LayoutPanel";
            p.AutoScroll = true;
            p.AutoSize = false;
            p.Dock = DockStyle.Fill;
            return p;
        }

        private void PopulateControlSetsToServiceTable(string key)
        {
            foreach (string filePath in services[key]) {
                ControlSet controlSet = new ControlSet(this, filePath);
                ControlSets.Add(controlSet);
                serviceTables[key] = AddNewRow(serviceTables[key], controlSet);
                Program.ExcelWriters[key].OpenOrCreateNewWorkBook(Directory.GetCurrentDirectory() + key + ".xlsx");
            }
        }

        private void BuildServiceTable(string key)
        {
            if (!serviceTables.ContainsKey(key)) {
                serviceTables.Add(key, NewTableLayoutPanel());
            }
            if (!Program.ExcelWriters.ContainsKey(key)) {
                Program.ExcelWriters.Add(key, new ExcelWriter());
            }
        }

        private TableLayoutPanel NewTableLayoutPanel()
        {
            TableLayoutPanel tableLayoutPanel = new TableLayoutPanel();
            tableLayoutPanel.AutoSize = true;
            tableLayoutPanel.GrowStyle = TableLayoutPanelGrowStyle.AddRows;
            tableLayoutPanel.BackColor = Color.Transparent;
            tableLayoutPanel.ColumnCount = 6;
            tableLayoutPanel.Location = new Point(0, 0);
            tableLayoutPanel.Name = "controlSet";
            tableLayoutPanel.Padding = new Padding(3);
            tableLayoutPanel.Dock = DockStyle.Top;
            tableLayoutPanel.AutoScroll = false;
            tableLayoutPanel.AutoScrollMinSize = new Size(0, 0);
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent));
            return tableLayoutPanel;
        }

        private void AddTab(string key)
        {
            bool found = false;
            foreach (TabPage tab in tabControl1.TabPages) {
                if (tab.Name.Equals(key)) {
                    found = true;
                }
            }
            if (!found)
                tabControl1.TabPages.Add(key, key);
        }

        private void AddFilesToServiceCollection(DragEventArgs e)
        {
            string[] filePaths = (string[])(e.Data.GetData(DataFormats.FileDrop));
            foreach (string filePath in filePaths) {
                if (filePath.EndsWith(".txt")) {
                    string fileName = FileUtils.GetFileName(filePath);
                    AddToServiceCollection(filePath, fileName);
                }
            }
        }

        private void AddToServiceCollection(string filePath, string fname)
        {
            if (!services.ContainsKey(fname)) {
                services.Add(fname, new List<string>());
            }
            services[fname].Add(filePath);
        }

        private TableLayoutPanel AddNewRow(TableLayoutPanel tableLayoutPanel, ControlSet controlSet)
        {

            tableLayoutPanel.RowCount++;
            tableLayoutPanel.Controls.Add(controlSet.filePath_Label, 0, tableLayoutPanel.RowCount);
            tableLayoutPanel.Controls.Add(controlSet.convert_ComboBox, 1, tableLayoutPanel.RowCount);
            tableLayoutPanel.Controls.Add(controlSet.start_Button, 2, tableLayoutPanel.RowCount);
            tableLayoutPanel.Controls.Add(controlSet.info_Label, 3, tableLayoutPanel.RowCount);
            tableLayoutPanel.Controls.Add(controlSet.progressBar, 4, tableLayoutPanel.RowCount);
            tableLayoutPanel.Controls.Add(controlSet.cancel_Button, 5, tableLayoutPanel.RowCount);
            return tableLayoutPanel;
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

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start(@Directory.GetCurrentDirectory());
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach(ExcelWriter ew in Program.ExcelWriters.Values) {
                ew.Cleanup();
            }
            foreach(ControlSet cs in ControlSets) {
                try {
                    cs.Dispose();
                } catch {

                }
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            new ExcelWriter().ExportToExcel();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
