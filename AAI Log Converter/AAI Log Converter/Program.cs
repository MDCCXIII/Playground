using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Log;
using System.Diagnostics;
using AAI_Log_Converter.ExcelInterop;

namespace AAI_Log_Converter
{
    static class Program
    {
        //internal static string FileName;

        public static Dictionary<string, ExcelWriter> ExcelWriters = new Dictionary<string, ExcelWriter>();

        public static List<string> FileNames = new List<string>();

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try {
                if (args.Length > 0) {

                } else {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new Form1());
                }
            }
            catch (Exception e) {
                Debug.WriteLine(e);
            }
        }
    }
}
