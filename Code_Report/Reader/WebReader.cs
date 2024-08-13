using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;

namespace Code_Report
{
    public class WebReader
    {
        private string _url;
        private string _currentDir;
        private string _previousLink;
        WebDriver _driver;
        WebClient _webClient;
        WebDriverWait wait;

        public WebReader(bool headless = false)
        {
            try 
            {
                System.IO.DirectoryInfo di = new DirectoryInfo(Path.Combine(Directory.GetCurrentDirectory(), "Edge"));
                DirectoryInfo[] dirs = di.GetDirectories();
                foreach (DirectoryInfo dir in dirs)
                {
                    if (dir != dirs.Last())
                    {
                        dir.Delete(true);
                    }    
                }
                updateVersion();
            }
            catch 
            { 
            }

            EdgeOptions option = new EdgeOptions();
            if (headless)
            {
                option.AddArgument("--headless");
            }
            option.AddArgument("--silent");
            option.AddArgument("--disable-gpu");
            option.AddArgument("--log-level=3");
            option.AddArgument("--disable-notifications");
            option.AddArgument("--disable-popup-blocking");
            option.AddArgument("--blink-settings=imagesEnabled=false");
            EdgeDriverService service = EdgeDriverService.CreateDefaultService();
            service.SuppressInitialDiagnosticInformation = true;
            service.HideCommandPromptWindow = true;
            this._driver = new EdgeDriver(service, option);
            this._driver.Manage().Window.Minimize();
            wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(480));
            _currentDir = Environment.CurrentDirectory;
        }
        ~WebReader()
        {
            try
            {
                this._driver.Close();
                this._driver.Quit();
            }
            catch (Exception ex) { }
        }

        public void closeDriver()
        {
            try
            {
                var tabs = _driver.WindowHandles;
                foreach (var tab in tabs)
                {
                    _driver.SwitchTo().Window(tab);
                    _driver.Close();
                }
            }
            catch (Exception ex) { }
            this._driver.Quit();
        }
        public void runUrl(string Url)
        {
            this._url = Url;
            if (Uri.IsWellFormedUriString(Url, UriKind.Absolute))
            {
                //_driver.SwitchTo().NewWindow(WindowType.Tab);
                //_driver.SwitchTo().Window(_driver.WindowHandles.Last());
                _driver.Navigate().GoToUrl(Url);
            }
        }

        public void closeAllTab()
        {
            var tabs = _driver.WindowHandles;

            foreach (var tab in tabs)
            {
                if (tabs[0] != tab)
                {
                    _driver.SwitchTo().Window(tab);
                    _driver.Close();
                }
            }
        }
        public void downloadFile(string Url, string fileName)
        {
            WebClient webClient = new WebClient();
            webClient.DownloadFile(new Uri(Url), fileName);
        }

        public List<IWebElement> extractTable(By itemSpecifier)
        {
            List<IWebElement> tableList = new List<IWebElement>();
            tableList = _driver.FindElements(itemSpecifier).ToList();
            //if (tableList != null && tableList.Count > 0)
            //{
            //    foreach (var row in tableList)
            //    {
            //        var resultCols = row.FindElements(By.TagName("td"));
            //    }
            //}
            return tableList;
        }

        public bool tableHasItem(By itemSpecifier)
        {
            bool isHasItem = false;
            return isHasItem;
        }

        private IWebElement WaitUntilVisible(IWebElement elementToBeDisplayed, int secondsTimeout = 10)  //By itemSpecifier
        {
            var wait = new WebDriverWait(_driver, new TimeSpan(0, 0, secondsTimeout));
            var element = wait.Until<IWebElement>(_driver =>
            {
                try
                {
                    //var elementToBeDisplayed = _driver.FindElement(itemSpecifier);
                    if (elementToBeDisplayed.Displayed)
                    {
                        return elementToBeDisplayed;
                    }
                    return null;
                }
                catch (StaleElementReferenceException)
                {
                    return null;
                }
                catch (NoSuchElementException)
                {
                    return null;
                }
            });
            return element;
        }

        private IWebElement WaitUntilVisibleBy(By itemSpecifier, int secondsTimeout = 10)  //By itemSpecifier
        {
            var wait = new WebDriverWait(_driver, new TimeSpan(0, 0, secondsTimeout));
            var element = wait.Until<IWebElement>(_driver =>
            {
                try
                {
                    var elementToBeDisplayed = _driver.FindElement(itemSpecifier);
                    if (elementToBeDisplayed.Displayed)
                    {
                        return elementToBeDisplayed;
                    }
                    return null;
                }
                catch (StaleElementReferenceException)
                {
                    return null;
                }
                catch (NoSuchElementException)
                {
                    return null;
                }
            });
            return element;
        }

        private bool checkElementByInterval(IWebElement element,
                                            string condition,
                                            double totalTime,
                                            double timeStep)
        {
            bool isExist = false;
            double timeCount = 0;
            int count = 0;
            while (element != null && timeCount < totalTime && count < 100)
            {
                if (String.Equals(element.Text.ToLower(), condition.ToLower()))
                {
                    Thread.Sleep((int)(timeStep * 1000));
                    timeCount += timeStep;
                    count++;
                }
                else
                {
                    isExist = true;
                    break;
                }
            }
            return isExist;
        }

        // Custom for Code Report Tracking

        /// <summary>
        /// Track Report for https://icc-es.org
        /// </summary>
        public string trackICCESReport(string Url, string reportNumber, string targetDir)
        {
            if (_url != Url)
            {
                runUrl(Url);
                _url = Url;
            }
            string outputText = "";
            IWebElement searchBox = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("report_display_name")));
            ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);arguments[0].value = '';", searchBox);
            //searchBox.Clear();
            searchBox.SendKeys(reportNumber);
            IWebElement searchBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"reportstabinner\"]/div[2]/form/div[6]/div/button[1]")));
            searchBtn.SendKeys(Keys.Enter);

            IWebElement listingTable = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("esr_report_listing")));
            IWebElement firstRow = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"esr_report_listing\"]/tbody")));

            int count = 1;
            while (count < 30)
            {
                if (!firstRow.Text.Contains("No Reports found") && firstRow.Text != "")
                {
                    break;
                }
                Thread.Sleep(100);
                count++;
            }

            while (firstRow.Text == _previousLink && !firstRow.Text.Contains("No Reports found"))
            {
                if (count > 60)
                {
                    break;
                }
                Thread.Sleep(100);
                count++;
            }

            if (firstRow.Text.Contains("No Reports found"))
            {
                outputText = "No Reports found";
            }
            else
            {
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"esr_report_listing\"]/tbody/tr/td[1]")));
                IWebElement report = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"esr_report_listing\"]/tbody/tr/td[1]")));
                string tabLink = report.FindElements(By.TagName("a"))[0].GetAttribute("href");//tds[0].FindElements(By.TagName("a"))[0].GetAttribute("href");
                string name = report.Text;
                _driver.SwitchTo().NewWindow(WindowType.Tab);
                _driver.SwitchTo().Window(_driver.WindowHandles.Last());
                _driver.Navigate().GoToUrl(tabLink);
                IWebElement downloadBtn = wait.Until(ExpectedConditions.ElementIsVisible(By.Id("download-link")));
                string pdfLink = downloadBtn.GetAttribute("href").ToString();
                outputText = pdfLink;
                //downloadFileAsync(pdfLink, _currentDir + "//Data//" + targetDir + "//" + name + ".pdf");
                _driver.Close();
                _driver.SwitchTo().Window(_driver.WindowHandles.First());
            }
            _previousLink = firstRow.Text;
            return outputText;
        }

        /// <summary>
        /// Track Report for https://www.iapmoes.org/
        /// </summary>
        public string trackIAPMOESReport(string Url, string reportNumber, string targetDir)
        {
            if (_url != Url)
            {
                runUrl(Url);
                _url = Url;
            }
            string outputText = "";
            IWebElement searchBox = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"keyword\"]")));
            searchBox.Clear();
            searchBox.SendKeys(reportNumber.Split('-')[1]);
            IWebElement searchBtn = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"main-content\"]/form/div/div[3]/button")));
            wait.Until(ExpectedConditions.ElementToBeClickable(searchBtn)).Click();
            //wait.Until(ExpectedConditions.ElementIsVisible(By.Id("searchTable")));
            IWebElement table = wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"searchTable\"]/tbody")));

            int count = 1;
            while (count<30)
            {
                if (!table.Text.Contains("No data available in table") && table.Text != "")
                {
                    break;
                }
                Thread.Sleep(100);
                count++;
            }

            if (table.Text.Contains("No data available in table"))
            {
                outputText = "No Reports found";
            }
            else 
            {
                IList<IWebElement> trCollection = table.FindElements(By.TagName("tr"));
                IList<IWebElement> tdCollection;
                foreach (IWebElement element in trCollection)
                {
                    tdCollection = element.FindElements(By.TagName("td"));
                    if (tdCollection[0].Text.Contains(reportNumber.Split('-')[1]))
                    {
                        string tabLink = tdCollection[0].FindElement(By.TagName("a")).GetAttribute("href");
                        string name = tdCollection[0].Text;
                        if (name.Contains("Cancelled"))
                        {
                            outputText = name;
                        }
                        else
                        {
                            //downloadFileAsync(tabLink, _currentDir + "//Data//" + targetDir + "//" + name + ".pdf");
                            outputText = tabLink;
                        }
                        break;
                    }
                }        
            }
            
            return outputText;
        }

        /// <summary>
        /// Track Report for https://www.drjcertification.org
        /// </summary>
        public void trackDRJCERTIFICATIONReport(string Url, string reportNumber, string targetDir)
        {
            if (_url != Url)
            {
                runUrl(Url);
                _url = Url;
            }
            IWebElement searchBox = WaitUntilVisible(_driver.FindElement(By.XPath("//*[@id=\"report number\"]/div/div/div/div/input[1]")), 20);
            searchBox.Clear();
            searchBox.SendKeys(reportNumber);
            IWebElement searchBtn = WaitUntilVisible(_driver.FindElement(By.Id("report number-item-0")), 10);
            if (searchBtn != null)
            {
                searchBtn.Click();
                WaitUntilVisibleBy(By.XPath("//*[@id=\"root\"]/div/div[2]/div[1]/table"), 20);
                List<IWebElement> resultTable = extractTable(By.XPath("//*[@id=\"root\"]/div/div[2]/div[1]/table/tbody"));
                if (resultTable.Count > 0)
                {
                    foreach (IWebElement row in resultTable)
                    {
                        var tds = row.FindElements(By.TagName("td"));
                        string pdfLink = tds[1].FindElements(By.TagName("a"))[0].GetAttribute("href");
                        Console.WriteLine(pdfLink);
                    }
                }
            }
        }
        public bool RemoteFileExists(string url)
        {
            try
            {
                HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
                request.Method = "HEAD";
                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                response.Close();
                return (response.StatusCode == HttpStatusCode.OK);
            }
            catch
            {
                return false;
            }
        }
        public void downloadFileAsync(string Url, string fileName)
        {
            _webClient = new WebClient();
            _webClient.DownloadFileCompleted += webClient_DownloadFileCompleted;
            _webClient.DownloadProgressChanged += webClient_DownloadProgressChanged;
            _webClient.DownloadFileAsync(new Uri(Url), fileName);
        }

        private void webClient_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            Console.WriteLine("Download finished");
        }

        private void webClient_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            Console.WriteLine("Downloading... Progress: {0} ({1} MC / {2} MB)", e.ProgressPercentage, e.BytesReceived, e.TotalBytesToReceive / 1024);
        }

        public void updateVersion()
        {
            new WebDriverManager.DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.MatchingBrowser);
        }
    }
}
