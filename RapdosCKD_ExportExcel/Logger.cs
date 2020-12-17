using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RapdosCKD_ExportExcel
{
    class Logger
    {
        /// <summary>
        /// creates log folder if it doesn't exists
        /// </summary>
        /// <param name="path"></param>
        public static void VerifyDir(string path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
            {
                dir.Create();
            }
        }
        /// <summary>
        /// Creates log file named after today's date and writes messages per line. 
        /// This method will take log string and log type char.
        /// </summary>
        /// <param name="log"></param>
        /// <param name="st"></param>
        public static void Write(string log,char st)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Log");
            VerifyDir(path);
            string filename = DateTime.Now.ToString("yyyyMMdd") + ".txt";
            string filepath = Path.Combine(path, filename);
            string state = "DEBUG";
            switch(st)
            {
                case 'd':
                   state = "DEBUG";
                   break;
                case 'e':
                    state = "ERROR";
                    break;
                case 'w':
                    state = "WARNING";
                    break;
                case 'n':
                    state = "";
                    break;
            }
            ///store in string builder first
            StringBuilder sb = new StringBuilder();
            sb.Append("\n[" + state + "]  " + log);
            ///and append the string builder as a string in the following text log file
            File.AppendAllText(filepath, sb.ToString());
            sb.Clear();
        }
    }
}
