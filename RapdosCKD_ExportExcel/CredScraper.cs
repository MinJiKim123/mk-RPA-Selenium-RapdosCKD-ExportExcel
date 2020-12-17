using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace RapdosCKD_ExportExcel
{
   
    class CredScraper
    {
        private string FILENAME;
        private string FILEPATH;


        public CredScraper()
        {
            FILENAME = "config.xml";
            FILEPATH = Path.Combine(Directory.GetCurrentDirectory(), FILENAME);
        }

        /// <summary>
        /// Gets the company information value from config.xml file. 
        /// </summary>
        /// <returns></returns>
        public List<Comp> Scrape()
        {           
            XDocument X = XDocument.Load(FILEPATH);           
            var comp = X.Element("config").Element("rapdos").Elements("company");            
            List <Comp> docs = (from doc in comp
                               select new Comp
                               {
                                   Code = (string)doc.Element("code"),
                                   Name = (string)doc.Element("name"),
                                   Username = (string)doc.Element("username"),
                                   Password = (string)doc.Element("password")
                               }).ToList();           
            return docs;
        }

       

    }
}
