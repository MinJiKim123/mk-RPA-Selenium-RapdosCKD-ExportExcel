using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RapdosCKD_ExportExcel
{
    public interface IWebScrape
    {
        void Execute();

        void SaveExl(IWebDriver driver, bool isTwice);

        void DeleteDoROMfiles(IWebDriver driver);

    }
}
