using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RapdosCKD_ExportExcel
{
    class Config
    {
        public static string URL;
        public static Browser DRIVER;
        public static string AnyERPPATH;

        private static string drivertemp;

        public static int INTERV_START;
        public static int INTERV_END;

        public static void GetVal()
        {
            
            string filename = "config.xml";
            string path = Path.Combine(Directory.GetCurrentDirectory(), filename);

            XDocument X = XDocument.Load(path);
            var pele = X.Element("config").Element("webconfig");

            URL = (string)pele.Element("url");
            string browserst = (string)pele.Element("browser");
            
            drivertemp = browserst.ToLower();
            switch (drivertemp)
            {
                case "chrome":
                    DRIVER = Browser.CHROME;
                    break;
                case "edge":
                    DRIVER = Browser.EDGE;
                    break;

                case "ie":
                    DRIVER = Browser.IE;
                    break;

                default:
                    DRIVER = Browser.CHROME;
                    Logger.Write("file not read", 'e');
                    break;


            }

            var pele2 = X.Element("config").Element("winconfig");
            AnyERPPATH = (string)pele2.Element("path");
           var dele =  X.Element("config").Element("rapdos").Element("date");
            INTERV_START = (int) dele.Element("start");
            INTERV_END = (int)dele.Element("end");

        }
        
        public static void ENDPROCESS()
        {
            Console.WriteLine("ending process");
            Process.GetCurrentProcess().CloseMainWindow();
            Process.GetCurrentProcess().Kill();
        }


    }

    //IE driver code not implemented.
    public enum Browser
    {
        CHROME, EDGE, IE
    }
}
