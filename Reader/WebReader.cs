using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;

namespace Reader
{
    public class WebReader : IDisposable
    {
        private string _url;
        private bool _isHeadless;
        private string _downloadDir;
        private string _driverExecutablePath;
        private IWebDriver _driver;
        private WebDriverWait _wait;

        // 0 = not running, 1 = running
        private int _runningFlag;

        public Action<string> Status;

        public string URL { get => _url; set => _url = value; }
        public bool IsHeadless { get => _isHeadless; set => _isHeadless = value; }
        public string DownloadDir { get => _downloadDir; set => _downloadDir = value; }

        /// <summary>
        /// Indicates whether driver + wait were created successfully.
        /// Useful for external callers to quickly check readiness for searches.
        /// </summary>
        public bool IsReady => _driver != null && _wait != null;

        /// <summary>
        /// Indicates whether the web driver is currently performing work (navigating / searching / loading).
        /// Controlled internally by StartRunning/StopRunning (thread-safe).
        /// </summary>
        public bool IsRunning => Interlocked.CompareExchange(ref _runningFlag, 0, 0) == 1;

        private void StartRunning(string reason = null)
        {
            var old = Interlocked.Exchange(ref _runningFlag, 1);
            if (old == 0)
            {
                Status?.Invoke($"Running=true{(string.IsNullOrWhiteSpace(reason) ? "" : " (" + reason + ")")}");
            }
        }

        private void StopRunning(string reason = null)
        {
            var old = Interlocked.Exchange(ref _runningFlag, 0);
            if (old == 1)
            {
                Status?.Invoke($"Running=false{(string.IsNullOrWhiteSpace(reason) ? "" : " (" + reason + ")")}");
            }
        }

        /// <summary>
        /// Lightweight constructor that avoids blocking WebDriverManager network calls on the UI thread.
        /// It tries to use an already available driver first. If none found, it schedules an async driver update
        /// (best-effort) and proceeds using the default service (PATH) so initialization is faster.
        /// </summary>
        public WebReader(string Url = "", bool headless = false)
        {
            _url = Url;
            _isHeadless = headless;

            // Try to locate an already-downloaded msedgedriver without blocking network calls.
            try
            {
                var downloaded = FindDownloadedDriverExecutable();
                if (!string.IsNullOrWhiteSpace(downloaded) && File.Exists(downloaded))
                {
                    _driverExecutablePath = downloaded;
                    Status?.Invoke($"Found msedgedriver at: {downloaded}");
                }
                else
                {
                    // Schedule background attempt to download/copy a driver next to exe.
                    Task.Run(() => UpdateDriverAsyncPlaceNextToExe(Status));
                    Status?.Invoke("No local msedgedriver found. Background driver update scheduled; using PATH/default if available.");
                }
            }
            catch (Exception ex)
            {
                Status?.Invoke($"Driver discovery failed: {ex.Message}. Will attempt default service.");
            }

            var options = new EdgeOptions();
            // Prefer Eager page load so navigation completes faster for many pages
            try
            {
                options.PageLoadStrategy = PageLoadStrategy.Eager;
            }
            catch { /* If older bindings don't support, ignore */ }

            if (_isHeadless)
            {
                try { options.AddArgument("--headless"); } catch { }
            }

            // Common speed / stability flags for headless/background execution
            var args = new[]
            {
                "--disable-gpu",
                "--disable-extensions",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-notifications",
                "--disable-popup-blocking",
                "--disable-background-networking",
                "--disable-background-timer-throttling",
                "--blink-settings=imagesEnabled=false",
                "--silent",
                "--log-level=3"
            };

            foreach (var a in args)
            {
                try { options.AddArgument(a); } catch { }
            }

            try
            {
                // Prefer non-interactive download preferences; ignore if not supported in current EdgeOptions impl.
                options.AddUserProfilePreference("download.default_directory", Environment.CurrentDirectory);
                options.AddUserProfilePreference("download.prompt_for_download", false);
                options.AddUserProfilePreference("download.directory_upgrade", true);
                options.AddUserProfilePreference("profile.default_content_setting_values.images", 2);
            }
            catch
            {
                // ignore when runtime doesn't support these prefs
            }

            try
            {
                EdgeDriverService service;

                if (!string.IsNullOrWhiteSpace(_driverExecutablePath) && File.Exists(_driverExecutablePath))
                {
                    var dir = Path.GetDirectoryName(_driverExecutablePath);
                    service = EdgeDriverService.CreateDefaultService(dir);
                    Status?.Invoke($"Using msedgedriver from configured path: {_driverExecutablePath}");
                }
                else
                {
                    // fallback to default (PATH) service
                    service = EdgeDriverService.CreateDefaultService();
                    Status?.Invoke("Using default msedgedriver service (will search PATH).");
                }

                service.SuppressInitialDiagnosticInformation = true;
                service.HideCommandPromptWindow = true;

                // Create driver with a reasonable command timeout; reduce very long timeouts to speed failures.
                _driver = new EdgeDriver(service, options, TimeSpan.FromSeconds(120));

                // Reduce implicit wait; prefer explicit waits for responsiveness.
                try { _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0); } catch { }

                // Limit page load timeout to something practical (overrides very long defaults)
                try { _driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60); } catch { }

                // Keep browser minimized so it doesn't have to render full UI in some environments.
                try { _driver.Manage().Window.Minimize(); } catch { }

                if (!string.IsNullOrWhiteSpace(_url) && Uri.IsWellFormedUriString(_url, UriKind.Absolute))
                {
                    StartRunning("initial navigation");
                    try
                    {
                        _driver.Navigate().GoToUrl(_url);
                        Status?.Invoke($"Navigated to initial URL: {_url}");
                    }
                    finally
                    {
                        StopRunning("initial navigation");
                    }
                }

                // Default explicit wait shortened to 60s for responsiveness (use explicit waits in methods)
                _wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(60));

                // Report ready
                Status?.Invoke($"WebReader initialized {(string.IsNullOrWhiteSpace(_url) ? "" : $"for {_url}")}. Ready={IsReady} Running={IsRunning}");
            }
            catch (Exception ex)
            {
                Status?.Invoke($"Failed to create EdgeDriver instance: {ex.Message}");
                _driver = null;
                _wait = null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposing) return;

            if (_driver != null)
            {
                try
                {
                    _driver.Quit();
                }
                catch { /* ignore shutdown errors */ }

                try
                {
                    _driver.Dispose();
                }
                catch { /* ignore dispose errors */ }

                _driver = null;
                _wait = null;

                // ensure running flag cleared
                StopRunning("disposed");
                Status?.Invoke("WebReader disposed. Running=false Ready=false");
            }
        }

        public void Close()
        {
            if (_driver == null) return;

            try
            {
                _driver.Quit();
            }
            catch { /* ignore */ }

            try
            {
                _driver.Dispose();
            }
            catch { /* ignore */ }

            _driver = null;
            _wait = null;

            StopRunning("closed");
            Status?.Invoke("WebReader closed. Running=false Ready=false");
        }

        private string FindDownloadedDriverExecutable()
        {
            try
            {
                var rootCandidates = new List<string>
                {
                    Environment.CurrentDirectory,
                    Path.Combine(Environment.CurrentDirectory, ".wdm"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".wdm"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".wdm")
                };

                var found = new List<string>();
                foreach (var root in rootCandidates.Where(Directory.Exists))
                {
                    try
                    {
                        var files = Directory.EnumerateFiles(root, "msedgedriver.exe", SearchOption.AllDirectories);
                        found.AddRange(files);
                    }
                    catch
                    {
                        // ignore permission/IO exceptions and continue
                    }
                }

                var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
                foreach (var p in pathEnv.Split(Path.PathSeparator).Where(p => !string.IsNullOrWhiteSpace(p)))
                {
                    try
                    {
                        var candidate = Path.Combine(p, "msedgedriver.exe");
                        if (File.Exists(candidate)) found.Add(candidate);
                    }
                    catch { }
                }

                if (!found.Any()) return null;

                var best = found
                    .Select(f => new { Path = f, Time = File.GetLastWriteTimeUtc(f) })
                    .OrderByDescending(x => x.Time)
                    .FirstOrDefault();

                return best?.Path;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Background-friendly driver update that will attempt to download the latest msedgedriver
        /// and copy it next to the exe. Provides status updates through the provided delegate.
        /// </summary>
        public static Task UpdateDriverAsyncPlaceNextToExe(Action<string> status = null)
        {
            return Task.Run(() =>
            {
                try
                {
                    status?.Invoke("Requesting latest EdgeDriver from WebDriverManager (async)...");
                    new DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.Latest);

                    // Try to locate the downloaded driver without instantiating WebReader (avoid recursion/heavy init).
                    var rootCandidates = new List<string>
                    {
                        Environment.CurrentDirectory,
                        Path.Combine(Environment.CurrentDirectory, ".wdm"),
                        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".wdm"),
                        Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".wdm")
                    };

                    var found = new List<string>();
                    foreach (var root in rootCandidates.Where(Directory.Exists))
                    {
                        try
                        {
                            var files = Directory.EnumerateFiles(root, "msedgedriver.exe", SearchOption.AllDirectories);
                            found.AddRange(files);
                        }
                        catch { }
                    }

                    var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
                    foreach (var p in pathEnv.Split(Path.PathSeparator).Where(p => !string.IsNullOrWhiteSpace(p)))
                    {
                        try
                        {
                            var candidate = Path.Combine(p, "msedgedriver.exe");
                            if (File.Exists(candidate)) found.Add(candidate);
                        }
                        catch { }
                    }

                    var downloaded = found
                        .Select(f => new { Path = f, Time = File.GetLastWriteTimeUtc(f) })
                        .OrderByDescending(x => x.Time)
                        .FirstOrDefault()?.Path;

                    if (string.IsNullOrWhiteSpace(downloaded))
                    {
                        status?.Invoke("Async: Could not locate downloaded msedgedriver.");
                        return;
                    }

                    var exeFolder = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar);
                    var destPath = Path.Combine(exeFolder, Path.GetFileName(downloaded));
                    try
                    {
                        File.Copy(downloaded, destPath, overwrite: true);
                        status?.Invoke($"Async: Copied msedgedriver to application folder: {destPath}");
                    }
                    catch (Exception copyEx)
                    {
                        status?.Invoke($"Async: Failed to copy msedgedriver to exe folder: {copyEx.Message}. Driver remains at: {downloaded}");
                    }
                }
                catch (Exception ex)
                {
                    status?.Invoke($"UpdateDriverAsyncPlaceNextToExe failed: {ex.Message}");
                }
            });
        }

        public void RunUrl(string url)
        {
            _url = url;
            if (_driver == null) throw new InvalidOperationException("Driver is not initialized.");
            if (Uri.IsWellFormedUriString(url, UriKind.Absolute))
            {
                StartRunning("navigate");
                try
                {
                    Status?.Invoke($"Navigating to {url}...");
                    _driver.Navigate().GoToUrl(url);
                    Status?.Invoke($"Navigated to {url}");
                }
                finally
                {
                    StopRunning("navigate");
                }
            }
        }

        private WebDriverWait CreateWait(int timeoutSeconds)
        {
            if (_driver == null) throw new InvalidOperationException("Driver is not initialized.");
            return new WebDriverWait(_driver, TimeSpan.FromSeconds(timeoutSeconds));
        }

        public IWebElement WaitUntilElementExists(By elementLocator, int timeout = 10)
        {
            try
            {
                var wait = CreateWait(timeout);
                return wait.Until(ExpectedConditions.ElementExists(elementLocator));
            }
            catch (WebDriverTimeoutException)
            {
                throw;
            }
        }

        public IWebElement WaitUntilElementVisible(By elementLocator, int timeout = 10)
        {
            try
            {
                var wait = CreateWait(timeout);
                return wait.Until(ExpectedConditions.ElementIsVisible(elementLocator));
            }
            catch (WebDriverTimeoutException)
            {
                throw;
            }
        }

        public IWebElement WaitUntilElementClickable(By elementLocator, int timeout = 10)
        {
            try
            {
                var wait = CreateWait(timeout);
                return wait.Until(ExpectedConditions.ElementToBeClickable(elementLocator));
            }
            catch (WebDriverTimeoutException)
            {
                throw;
            }
        }

        //public void RunICCESWeb(ref Codes code)
        //{
        //    try
        //    {
        //        StartRunning("ICC-ES search");
        //        Status?.Invoke($"ICC-ES: Searching for {code.Number}...");
        //        if (_wait == null) _wait = CreateWait(30);

        //        var searchBox = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("report_display_name")));
        //        ((IJavaScriptExecutor)_driver).ExecuteScript("arguments[0].scrollIntoView(true);arguments[0].value = '';", searchBox);
        //        searchBox.Clear();
        //        searchBox.SendKeys(code.Number);

        //        var searchBtn = _wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"reportstabinner\"]/div[2]/form/div[6]/div/button[1]")));
        //        searchBtn.SendKeys(Keys.Enter);

        //        var listingTable = _wait.Until(ExpectedConditions.ElementIsVisible(By.Id("esr_report_listing")));
        //        var firstRow = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"esr_report_listing\"]/tbody")));

        //        string pdfLink = "No Reports found";

        //        // wait up to 3 seconds, polling 100ms
        //        const int attempts = 30;
        //        for (int i = 0; i < attempts; i++)
        //        {
        //            if (firstRow.Text.Contains(code.Number.Split('-').ElementAtOrDefault(1) ?? string.Empty))
        //            {
        //                var downloadBtn = _driver.FindElement(By.XPath("//*[@id=\"esr_report_listing\"]/tbody/tr/td[1]/a[1]"));
        //                var tmpTxt = downloadBtn.GetDomProperty("href")?.ToString() ?? string.Empty;
        //                pdfLink = $"https://cdn-v2.icc-es.org/wp-content/uploads/report-directory/{code.Number}.pdf";
        //                break;
        //            }

        //            Thread.Sleep(100);
        //        }

        //        code.Link = pdfLink;
        //        code.HasCheck = true;
        //        code.LastCheck = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        //        Status?.Invoke($"ICC-ES: Search completed for {code.Number}. Link: {code.Link}");
        //    }
        //    catch (Exception ex)
        //    {
        //        Status?.Invoke($"ICC-ES: Search failed for {code.Number}: {ex.Message}");
        //        throw;
        //    }
        //    finally
        //    {
        //        StopRunning("ICC-ES search");
        //    }
        //}

        //public void RunIAPMOWeb(ref Codes code)
        //{
        //    try
        //    {
        //        StartRunning("IAPMO search");
        //        Status?.Invoke($"IAPMO: Searching for {code.Number}...");
        //        if (_wait == null) _wait = CreateWait(30);

        //        var reg = new Regex(@"[0-9]+");
        //        var searchBox = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"ctl00_pageContent_UES_txtSearch\"]")));
        //        searchBox.Clear();
        //        searchBox.SendKeys(reg.Match(code.Number).Value);

        //        var searchBtn = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"ctl00_pageContent_UES_btnAny\"]")));
        //        searchBtn.SendKeys(Keys.Enter);

        //        var statusText = _wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id=\"aspnetForm\"]")));

        //        string pdfLink = "No Reports found";

        //        const int attempts = 30;
        //        for (int i = 0; i < attempts; i++)
        //        {
        //            if (statusText.Text.Contains("All Evaluation Reports displayed are currently valid and active"))
        //            {
        //                var downloadBtn = _driver.FindElement(By.XPath("//*[@id=\"ctl00_pageContent_UES_gvReports\"]/tbody/tr[2]/td[1]/a"));
        //                var tmpTxt = downloadBtn.GetDomProperty("href")?.ToString() ?? string.Empty;
        //                if (tmpTxt.Contains(reg.Match(code.Number).Value)) pdfLink = tmpTxt;
        //                break;
        //            }

        //            Thread.Sleep(100);
        //        }

        //        code.Link = pdfLink;
        //        code.HasCheck = true;
        //        code.LastCheck = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

        //        Status?.Invoke($"IAPMO: Search completed for {code.Number}. Link: {code.Link}");
        //    }
        //    catch (Exception ex)
        //    {
        //        Status?.Invoke($"IAPMO: Search failed for {code.Number}: {ex.Message}");
        //        throw;
        //    }
        //    finally
        //    {
        //        StopRunning("IAPMO search");
        //    }
        //}
    }
}