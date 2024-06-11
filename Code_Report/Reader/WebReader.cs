using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;

namespace Code_Report
{
    public class WebReader
    {
        private string _url;
        private string _currentDir;
        WebDriver _driver;
        WebDriverWait _wait;
        WebClient _webClient;

        public WebReader()
        {
            EdgeOptions option = new EdgeOptions();
            //option.AddArgument("--headless");
            option.AddArgument("--silent");
            option.AddArgument("--disable-gpu");
            option.AddArgument("--log-level=3");
            //option.AddArgument("--remote-debugging-port=9222");
            EdgeDriverService service = EdgeDriverService.CreateDefaultService();
            service.SuppressInitialDiagnosticInformation = true;
            service.HideCommandPromptWindow = true;

            this._driver = new EdgeDriver(service, option);
            //this._driver.Manage().Window.Minimize();
            WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(10));

            _currentDir = Environment.CurrentDirectory;
        }
        ~WebReader()
        {
            //this._driver.Close();
            //this._driver.Quit();
        }

        public void runUrl(string Url)
        {
            this._url = Url;
            if (Uri.IsWellFormedUriString(Url, UriKind.Absolute))
            {
                _driver.Navigate().GoToUrl(Url);
            }
        }

        public void closeDriver()
        {
            this._driver.Close();
            this._driver.Quit();
        }

        public void closeAllTab()
        {
            var tabs = _driver.WindowHandles;

            foreach (var tab in tabs)
            {
                // "tab" is a string like "CDwindow-6E793DA3E15E2AB5D6AE36A05344C68"
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

        public void downloadFileAsync(string Url, string fileName)
        {
            _webClient = new WebClient();
            _webClient.DownloadFileCompleted += webClient_DownloadFileCompleted;
            _webClient.DownloadProgressChanged += webClient_DownloadProgressChanged;
            _webClient.DownloadFileAsync(new Uri(Url), fileName);

        }

        private void _webClient_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void webClient_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            Console.WriteLine("Download finished");
        }

        private void webClient_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            Console.WriteLine("Downloading... Progress: {0} ({1} bytes / {2} bytes)", e.ProgressPercentage, e.BytesReceived, e.TotalBytesToReceive);
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
        public string trackICCESReport(string Url, string reportNumber)
        {
            if (_url != Url)
            {
                runUrl(Url);
                _url = Url;
            }
            string outputText = "";
            IWebElement searchBox = _driver.FindElement(By.Id("report_display_name"));
            searchBox.Clear();
            searchBox.SendKeys(reportNumber);
            IWebElement searchBtn = _driver.FindElement(By.XPath("//*[@id=\"reportstabinner\"]/div[2]/form/div[6]/div/button[1]"));
            searchBtn.Click();

            WaitUntilVisible(_driver.FindElement(By.Id("esr_report_listing")), 60);
            bool hasReport = checkElementByInterval(_driver.FindElement(By.XPath("//*[@id=\"esr_report_listing\"]/tbody")), "No Reports found", 10, 0.1);

            List<IWebElement> resultTable = extractTable(By.Id("esr_report_listing"));
            if (resultTable.Count > 0 && hasReport)
            {
                foreach (IWebElement row in resultTable)
                {
                    string cellText = row.Text;
                    var tds = row.FindElements(By.TagName("td"));
                    string tabLink = tds[0].FindElements(By.TagName("a"))[0].GetAttribute("href");
                    string name = tds[0].Text;
                    _driver.SwitchTo().NewWindow(WindowType.Tab);
                    _driver.SwitchTo().Window(_driver.WindowHandles.Last());
                    _driver.Navigate().GoToUrl(tabLink);
                    //_driver.Manage().Window.Minimize();
                    IWebElement downloadBtn = WaitUntilVisible(_driver.FindElement(By.Id("download-link")), 10);
                    string pdfLink = downloadBtn.GetAttribute("href");
                    Console.WriteLine("Start Download File");
                    downloadFileAsync(pdfLink, _currentDir + "//" + name + ".pdf");

                    _driver.Close();
                    _driver.SwitchTo().Window(_driver.WindowHandles.First());
                    outputText = pdfLink;
                }
            }
            else
            {
                outputText = "Code: " + reportNumber + " has no report!";
            }
            return outputText;
        }

        /// <summary>
        /// Track Report for https://www.iapmoes.org/
        /// </summary>
        public string trackIAPMOESReport(string Url, string reportNumber)
        {
            if (_url != Url)
            {
                runUrl(Url);
                _url = Url;
            }
            string outputText = "";
            IWebElement searchBox = _driver.FindElement(By.XPath("//*[@id=\"keyword\"]"));
            searchBox.Clear();
            searchBox.SendKeys(reportNumber);
            IWebElement searchBtn = _driver.FindElement(By.XPath("//*[@id=\"main-content\"]/form/div/div[3]/button"));
            searchBtn.Click();

            WaitUntilVisible(_driver.FindElement(By.Id("searchTable")), 60);
            bool hasReport = checkElementByInterval(_driver.FindElement(By.XPath("//*[@id=\"searchTable\"]/tbody")), "No data available in table", 10, 0.1);

            List<IWebElement> resultTable = extractTable(By.Id("searchTable"));
            if (resultTable.Count > 0 && hasReport)
            {
                foreach (IWebElement row in resultTable)
                {
                    string cellText = row.Text;
                    var tds = row.FindElements(By.TagName("td"));
                    string tabLink = tds[0].FindElements(By.TagName("a"))[0].GetAttribute("href");
                    string name = tds[0].Text;
                    Console.WriteLine("Start Download File");
                    downloadFileAsync(tabLink, _currentDir + "//" + name + ".pdf");
                    outputText = tabLink;
                }
            }
            else
            {
                outputText = "Code: " + reportNumber + " has no report!";
            }
            return outputText;
        }

        /// <summary>
        /// Track Report for https://www.drjcertification.org
        /// </summary>
        public void trackDRJCERTIFICATIONReport(string Url, string reportNumber)
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
    }
}
