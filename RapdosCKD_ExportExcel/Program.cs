using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RapdosCKD_ExportExcel
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Logger.Write("================[" + DateTime.Now.ToString("dddd, dd MMMM yyyy HH:mm:ss") + "]================",'n');
            Config.GetVal();
            CredScraper cs = new CredScraper();
            IWebScrape ws;
            foreach (Comp cred in cs.Scrape())
            {
                ws = new CWebScraper(cred, Config.DRIVER);
                ws.Execute();
            }
             
            
            Logger.Write("프로세스를 종료합니다", 'd');
            Config.ENDPROCESS();
            
           
        }


    }
}
