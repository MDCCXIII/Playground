using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AAI_Log_Converter.FileIO
{
    public static class FileUtils
    {

        public static List<string> FileToList(string filePath)
        {
            List<string> result = new List<string>();
            using (TextReader tr = new StreamReader(filePath)) {
                string line;
                while ((line = tr.ReadLine()) != null) {
                    result.Add(line);
                }
            }
            return result;
        }

        public static int GetFileLineCount(string filePath)
        {
            return File.ReadAllLines(filePath).Count();
        }

        public static bool IsFileinUse(string filePath)
        {
            bool result = false;
            Stream s = null;
            try {
                if (FileExists(filePath)) {
                    s = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
            }
            catch (IOException) {
                result = true;
                //TODO: Add Logging
                //MessageBox.Show("The File " + file + " is currently open.\nPlease close the file and try again.");
            }
            finally {
                if (s != null) {
                    s.Close();
                }
            }
            return result;
        }

        private static void DeleteFile(string filePath)
        {
            File.Delete(filePath);
        }

        public static bool FileExists(string filePath)
        {
            return File.Exists(filePath);
        }

        public static string GetFileName(string filePath)
        {
            if (filePath.Contains("_")) {
                filePath = filePath.Split('_')[1];
            }
            return Path.GetFileName(filePath).Split('.')[0];
        }
    }
}
