using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RapdosCKD_ExportExcel
{
    class CWebScraper : IWebScrape
    {
        private readonly string userid;
        private readonly string password;
        private readonly string compname;
        private readonly string url = Config.URL;
        private int count = 0;
        private IWebDriver driver;
        

        public CWebScraper(Comp cred,Browser br)
        {
            userid = cred.Username;
            password = cred.Password;
            compname = cred.Name;
   
            driver=getDriver(br);
        }

        public void Execute()
        {
            count++;
            Logger.Write("사용자ID : " + userid + "로 Rapdos 를 실행합니다.", 'd');
           
            
            try
            {
                driver.Url = url;
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                
                ////////////////////////////////LOGIN////////////////////////////////////
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("Userid")));
                IWebElement wuserId = driver.FindElement(By.Name("Userid"));
                IWebElement wpassword = driver.FindElement(By.Name("Passwd"));
                IWebElement loginBtn = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td/input"));

                wuserId.Clear();
                wuserId.SendKeys(userid);
                wpassword.SendKeys(password);
                loginBtn.Click();               

                if (!isAlertPresent(driver))
                {                  
                    Console.WriteLine("no duplicate user");
                }
                else
                {
                    
                    IAlert alert = driver.SwitchTo().Alert();
                    if (alert.Text.Contains("비밀번호가 틀렸습니다"))
                    {
                        Console.WriteLine("wrong password..");
                        Logger.Write("비밀번호가 올바르지 않습니다. 사용자 ID : " + userid + " 에 대하여 자동화 작업을 중단합니다.", 'w');
                        //count = 2;
                        driver.Quit();
                        return;
                    }else
                    {
                        
                        alert.Accept();
                        Logger.Write("ID : " + userid + " 를 사용중인 다른 사용자 밀어내고 로그인 시작합니다", 'd');
                        Console.WriteLine("kicking out duplicate user and re-logging in");
                        IWebElement _password = driver.FindElement(By.Name("Passwd"));
                        IWebElement _loginBtn = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td/input"));
                        _password.SendKeys(password);
                        _loginBtn.Click();
                        //Thread.Sleep(500);
                        if(isAlertPresent(driver))
                        {
                            IAlert alert2 = driver.SwitchTo().Alert();
                            if (alert2.Text.Contains("비밀번호가 틀렸습니다"))
                            {
                                Console.WriteLine("wrong password..");
                                Logger.Write("비밀번호가 올바르지 않습니다. 사용자 ID : " + userid + " 에 대하여 자동화 작업을 중단합니다.", 'w');
                                //count = 2;
                                driver.Quit();
                                return;
                            }
                        }                                             
                        Thread.Sleep(900);
                    }
                    

                }
                ///작은 화면 : 사용자 비밀번호 재차 확인
                if (driver.WindowHandles.Count() > 1)
                {
                    String currentWindow = driver.CurrentWindowHandle;
                    foreach (String winHandle in driver.WindowHandles)
                    { driver.SwitchTo().Window(winHandle); }
                    Console.WriteLine(driver.Url);

                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("Passwd")));
                    IWebElement _password2 = driver.FindElement(By.Name("Passwd"));
                    IWebElement _loginBtn2 = driver.FindElement(By.XPath("/html/body/form/img"));

                    _password2.SendKeys(password);
                    _loginBtn2.Click();
                    //this part needs to be organized into single method. too messy
                    if (!isAlertPresent(driver))
                    {
                        Console.WriteLine("No alert");
                    }
                    else
                    {
                        IAlert alert2 = driver.SwitchTo().Alert();
                        alert2.Accept();
                        IWebElement _password2_2 = driver.FindElement(By.Name("Passwd"));
                        IWebElement _loginBtn2_2 = driver.FindElement(By.XPath("/html/body/form/img"));

                        _password2_2.SendKeys(password);
                        _loginBtn2_2.Click();

                    }
                    driver.SwitchTo().Window(currentWindow);
                    Console.WriteLine(driver.Url);
                }

                Logger.Write("로그인 완료. CKD 미납현황조회를 시작합니다.", 'd');
                ////////////////////////////////CKD 미닙현황조회////////////////////////////////////
                driver.SwitchTo().Frame("menu");
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"toptbl\"]/tbody/tr[2]/td/table/tbody/tr[9]/td[2]/table/tbody/tr[1]/td[1]/a")));
                IWebElement navigate1 = driver.FindElement(By.XPath("//*[@id=\"toptbl\"]/tbody/tr[2]/td/table/tbody/tr[9]/td[2]/table/tbody/tr[1]/td[1]/a"));
                navigate1.Click();
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/a")));
                IWebElement navigate1_1 = driver.FindElement(By.XPath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/a"));
                navigate1_1.Click();

                
                driver.SwitchTo().DefaultContent();
                Thread.Sleep(600);
                driver.SwitchTo().Frame("work");
                Thread.Sleep(600);

                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("vnt1")));
                
                ///회사구분
                IWebElement filter2 = driver.FindElement(By.Name("comp"));
                ///공장
                IWebElement filter3 = driver.FindElement(By.Name("plnt"));
                ///납기일자start
                IWebElement filter4 = driver.FindElement(By.Name("nidt_f"));
                ///납기일자end
                IWebElement filter5 = driver.FindElement(By.Name("nidt_t"));

                ///get current month,year
                DateTime now = DateTime.Now;

                ///get current month +1 -1
                DateTime start_temp = new DateTime(now.Year, now.Month, 1);
                DateTime start = start_temp.AddMonths(Config.INTERV_START*-1);
                DateTime end = start_temp.AddMonths(1 + Config.INTERV_END).AddDays(-1);

                ///re-arrange the format into yyyymmdd
                string cmonthstart = start.ToString("yyyyMMdd");
                string cmonthend = end.ToString("yyyyMMdd");

                
                var selectElement2 = new SelectElement(filter2);
                filter2.Click();
                selectElement2.SelectByValue("EHMC");

                var selectElement3 = new SelectElement(filter3);
                filter3.Click();
                selectElement3.SelectByValue("HK11");

                filter4.Clear();
                filter4.SendKeys(cmonthstart);

                filter5.Clear();
                filter5.SendKeys(cmonthend);

                //load data
                IWebElement loadBtn = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[1]/td[3]/img[1]"));
                loadBtn.Click();
                Thread.Sleep(3000);

                DeleteDoROMfiles(driver);

                ////////////////////////////////엑셀파일 호출 및 저장////////////////////////////////////
                IWebElement exportBtn = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[1]/td[3]/img[2]"));
                Logger.Write("엑셀파일 다운로드 및 저장을 실행합니다...", 'd'); 
                exportBtn.Click();
                Thread.Sleep(6000);

              
                bool isDownloaded = checkDownloadComplete();

                //after the download
                if(isDownloaded)
                    SaveExl(driver,false);
                else
                {
                    Logger.Write("엑셀파일 다운로드 오류. 해당 사용자에대하여 조회 자동화를 중단합니다.사용자 ID : " + userid + " 에 대하여 수동 작업하십시오.", 'e');
                    driver.Quit();
                    return;

                }
                   

                ///동국 실업 2차 조회
                if(compname.Contains("동국실업"))
                {
                    Logger.Write("동국실업 T1 업체 U09F 에 대해 조회 시작합니다...",'d');

                    ///goes to 'menu' frame for refreshing of all the content in the 'work' frame
                    driver.SwitchTo().DefaultContent();
                    Thread.Sleep(600);
                    driver.SwitchTo().Frame("menu");
                    Thread.Sleep(600);
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/a")));
                    IWebElement navigate1_1_2 = driver.FindElement(By.XPath("/html/body/table/tbody/tr[3]/td/table/tbody/tr[2]/td/table/tbody/tr[10]/td/table/tbody/tr[3]/td[2]/table/tbody/tr[1]/td[1]/a"));
                    navigate1_1_2.Click();

                    ///back to the work frame
                    driver.SwitchTo().DefaultContent();
                    Thread.Sleep(600);
                    driver.SwitchTo().Frame("work");
                    Thread.Sleep(600);

                    ///repeat all the work with different condition
                    wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Name("vnt1")));
                    IWebElement filter1 = driver.FindElement(By.Name("vnt1"));
                    IWebElement filter2_1 = driver.FindElement(By.Name("comp"));
                    IWebElement filter3_1 = driver.FindElement(By.Name("plnt"));
                    IWebElement filter4_1 = driver.FindElement(By.Name("nidt_f"));
                    IWebElement filter5_1 = driver.FindElement(By.Name("nidt_t"));
                    var selectElement1 = new SelectElement(filter1);
                    filter1.Click();
                    selectElement1.SelectByValue("UO9F");

                    var selectElement2_1 = new SelectElement(filter2_1);
                    filter2_1.Click();
                    selectElement2_1.SelectByValue("EHMC");

                    var selectElement3_1 = new SelectElement(filter3_1);
                    filter3_1.Click();
                    selectElement3_1.SelectByValue("HK11");

                    filter4_1.Clear();
                    filter4_1.SendKeys(cmonthstart);

                    filter5_1.Clear();
                    filter5_1.SendKeys(cmonthend);

                    
                    IWebElement loadBtn2 = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[1]/td[3]/img[1]"));
                    loadBtn2.Click();
                    Thread.Sleep(3000);
                    DeleteDoROMfiles(driver);
                    IWebElement exportBtn2 = driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[1]/td/table/tbody/tr[1]/td[3]/img[2]"));
                    exportBtn2.Click();
                    Logger.Write("동국실업_U09F 엑셀 파일 호출 시작합니다..", 'd');                  

                    bool isDownloaded_2 = checkDownloadComplete();
                    if (isDownloaded)
                        SaveExl(driver, true);
                    else
                    {
                        Logger.Write("엑셀파일 다운로드 오류. 해당 사용자에대하여 조회 자동화를 중단합니다.사용자 ID : " + userid + " 에 대하여 수동 작업하십시오.", 'e');
                        driver.Quit();
                        return;

                    }

                }
                

                driver.SwitchTo().DefaultContent();
                Thread.Sleep(600);
                driver.SwitchTo().Frame("up");
                Thread.Sleep(600);
                IWebElement logoutBtn = driver.FindElement(By.XPath("/html/body/div[1]/img[2]"));
                logoutBtn.Click();
                
                driver.Quit();
            } catch(Exception ex)
            {
                Logger.Write("웹 자동화 실행중 오류; Method : Execute(); 오류정보 : " + lineNumber(ex) + ">>" + ex, 'e');
                Console.WriteLine("Error while executing webscraper : " + ex);
                ///첫번째 오류시 재실행 시킨다
                if (count == 1)
                {
                    Logger.Write("사용자 ID : " + userid + " 로 조회를 재시작합니다...", 'd');
                    Execute();
                }else///2번째 오류는 프로세스를 중단하고, 다음 차례로 넘긴다.
                {
                    Logger.Write("해당 사용자에대하여 조회 자동화를 중단합니다.사용자 ID : " + userid + " 에 대하여 수동 작업하십시오.", 'w');
                    driver.Quit();
                    return;
                    
                }
                
            }          
            
        }

        private bool isAlertPresent(IWebDriver driver)
        {         
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch (Exception ex)
            {
                
                return false;
            }           
        }

        ///This method copies the ROM3W.xls file from Download folder into folder 
        ///named after company name taken from config.txt file. 
        public void SaveExl(IWebDriver driver, bool isTwice)
        {
            try
            {
                DateTime now = DateTime.Now;
                string stnow = now.ToString("yyyyMMdd");
                string dirname = compname;
                string dfilename = "ROM3W.xls";
                string filename = dfilename;
                if(compname.Contains("동국실업"))
                {
                    if (isTwice)
                        filename = "ROM3W_U09F.xls";
                    else
                        filename = "ROM3W_U678.xls";
                }
                string newfilename = stnow + "_" + filename;
                string dpath = KnownFolders.GetPath(KnownFolder.Downloads);
                string[] dfilepaths = Directory.GetFiles(dpath, dfilename);
                string dfilepath = dfilepaths.First();


                //create directory
                string root = @"C:\RAPDOSEXL";
                string finaldir = Path.Combine(root, dirname);
                string finalpath = Path.Combine(finaldir, newfilename);
                if (!Directory.Exists(root))
                {
                    Directory.CreateDirectory(root);
                }
                if (!Directory.Exists(finaldir))
                {
                    Directory.CreateDirectory(finaldir);
                }
                Console.WriteLine("file to copy : " + dfilepath);
                Console.WriteLine("file after copy : " + finalpath);
                File.Copy(dfilepath, finalpath, true);
                Logger.Write(compname + " 엑셀파일 저장 완료; 파일경로 : " + finalpath, 'd');
            }
            catch (Exception ex)
            {
                Logger.Write("엑셀파일 복사중 오류; method : saveExl; 오류정보 : " + lineNumber(ex) + ">>" + ex, 'e');
                Console.WriteLine("Error while copying excel file : " + ex);
                driver.Quit();
                Config.ENDPROCESS();
            }


        }

        ///This method is to make sure all the ROM3W.xls file is deleted before the new file downloaded for another company. 
        ///It was created because the file name was same for all the company (as ROM3W) and by deleting all the same named file in 
        ///Download folder, it will reduce the complication and any possible error of name matching later, when copying to new folder
        public void DeleteDoROMfiles(IWebDriver driver)
        {
            try
            {
                string dpath = KnownFolders.GetPath(KnownFolder.Downloads);
                string[] dfilepaths = Directory.GetFiles(dpath, "ROM3W*" + ".xls");
    
                if (dfilepaths.Length > 0)
                {
                    foreach (string f in dfilepaths)
                    {
                        File.Delete(f);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Write("파일을 지우는중 오류.사용자 ID : "+ userid+" 엑셀파일 수동 다운로드 하십시오.; method : deleteDoROMfiles; 오류정보 : " + ex, 'e');
                Console.WriteLine("Error while deleting ROM files : " + ex);
                driver.Quit();
                Config.ENDPROCESS();
            }
        }

        ///This method checks whether the file download is complete in the Download folder
        private bool checkDownloadComplete()
        {
            int count = 0;
            string dpath = KnownFolders.GetPath(KnownFolder.Downloads);
            string fp = Path.Combine(dpath, "ROM3W.xls");
            Console.WriteLine(fp);
            Console.WriteLine(File.Exists(fp));
            do
            {
                ++count;
                Thread.Sleep(900);
            } while (!File.Exists(fp) && count < 30);
            if (count == 30)
                return false;
            return true;
        }
        private string lineNumber(Exception e)
        {

            int linenum = 0;
            try
            {               
                linenum = Convert.ToInt32(e.StackTrace.Substring(e.StackTrace.LastIndexOf(' ')));
                
            }
            catch
            {
                //Stack trace is not available
            }
            return "Line : " + linenum;
        }

        /// Initiates driver regarding to the config file browser variable
        private IWebDriver getDriver(Browser br)
        {
            try
            {
                switch (br)
                {
                    case Browser.CHROME:
                        //ChromeOptions co = new ChromeOptions();
                        //co.AddArgument("headless");
                        return new ChromeDriver();
                    case Browser.EDGE:
                        string edriverPath = Directory.GetCurrentDirectory();
                        return new EdgeDriver();

                    default:
                        return new ChromeDriver();
                }
            }
            catch (Exception ex)
            {

                Logger.Write("Driver error; Error info : " + lineNumber(ex) + ">>" + ex, 'e');
                //driver.Quit();          
                Config.ENDPROCESS();
                return null;
                

            }
        }

    }
}
