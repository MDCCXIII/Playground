using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace TextToExcelConverter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            try {
                Application.Run(new Form1());
            } catch (Exception ex) {
                Debug.WriteLine(ex);
            }
        }
    }
}
