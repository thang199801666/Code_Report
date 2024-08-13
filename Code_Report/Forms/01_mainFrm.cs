using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Code_Report
{
    public partial class mainFrm : Form
    {
        public string _currentDirectory = Directory.GetCurrentDirectory();
        public string _currentTab;
        WebReader _iccesReader;
        WebReader _iapmoesESWeb;
        Dictionary<string, int> _downloadObject;
        TabPage _currentPage;
        bool _allowChange = true;
        Thread _webThread;
        ExcelReader _xlExport;

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            System.Windows.Forms.Label label = new System.Windows.Forms.Label();
            System.Windows.Forms.TextBox textBox = new System.Windows.Forms.TextBox();
            System.Windows.Forms.Button buttonOk = new System.Windows.Forms.Button();
            System.Windows.Forms.Button buttonCancel = new System.Windows.Forms.Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new System.Windows.Forms.Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            DialogResult dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }
        public mainFrm()
        {
            InitializeComponent();
        }
        private void mainFrm_Load(object sender, EventArgs e)
        {
            string targetDir = Directory.GetCurrentDirectory() + "//Data//";
            if (!Directory.Exists(targetDir))
            {
                System.IO.Directory.CreateDirectory(targetDir);
            }
            string settingPath = Directory.GetCurrentDirectory() + "//Data//Settings.bin";
            if (File.Exists(settingPath))
            {
            WindowSetting setting = IOData.ReadFromBinaryFile<WindowSetting>(settingPath);
            btnLinkUpdated.BackColor = IOData.ListToColor(setting.FirstBtnColor);
            btnLinkNotFound.BackColor = IOData.ListToColor(setting.SecondBtnColor);
            btnLinkNotValid.BackColor = IOData.ListToColor(setting.ThirdBtnColor); 
            btnLinkCancelled.BackColor = IOData.ListToColor(setting.FourthBtnColor); 
            btnLinkCorrect.BackColor = IOData.ListToColor(setting.FifthBtnColor);
            }
            string binPath = Directory.GetCurrentDirectory() + "//Data//SST ESRs & ERs.bin";
            if (File.Exists(binPath))
            {
                Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
                foreach (var code in codeReports)
                {
                    DataGridViewRow row = (DataGridViewRow)mainTableView.Rows[0].Clone();
                    row.Cells[0].ToolTipText = code.Value.Link;
                    row.Cells[0].Value = code.Value.Number;
                    row.Cells[1].Value = code.Value.ProductCategory;
                    row.Cells[2].Value = code.Value.Description;
                    row.Cells[3].Value = code.Value.ProductsListed;
                    row.Cells[4].Value = code.Value.LatestCode;
                    row.Cells[5].Value = code.Value.IssueDate;
                    row.Cells[6].Value = code.Value.ExpirationDate;
                    row.Cells[7].Value = code.Value.WebType;
                    row.Cells[8].ToolTipText = @"Check Web      : Not Yet" + Environment.NewLine + "Check Update : Not Yet";
                    mainTableView.Rows.Add(row);
                }
            }
            mainTableView.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            _webThread = new Thread(callWebDriver);
            _webThread.Start();
        }
       
        private void mainFrm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                dynamic result = MessageBox.Show("Do You Want To Close Program?", "Code Report", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    if (_iccesReader != null) 
                    {
                        _iccesReader.closeDriver();  
                    }
                    if (_iapmoesESWeb != null)
                    {
                        _iapmoesESWeb.closeDriver();
                    }
                    System.Windows.Forms.Application.Exit();
                }

                if (result == DialogResult.No)
                {
                    e.Cancel = true;
                }
            }
            else //(e.CloseReason == CloseReason.WindowsShutDown)
            {
                //_iccesReader.closeDriver();
            }
            string settingPath = Directory.GetCurrentDirectory() + "//Data//Settings.bin";
            WindowSetting setting = new WindowSetting();    
            setting.FirstBtnColor = IOData.ColorToList(btnLinkUpdated.BackColor);
            setting.SecondBtnColor = IOData.ColorToList(btnLinkNotFound.BackColor);
            setting.ThirdBtnColor = IOData.ColorToList(btnLinkNotValid.BackColor);
            setting.FourthBtnColor = IOData.ColorToList(btnLinkCancelled.BackColor);
            setting.FifthBtnColor = IOData.ColorToList(btnLinkCorrect.BackColor);
            IOData.WriteToBinaryFile(settingPath, setting);
        }
       
        private void mainFrm_ResizeBegin(object sender, EventArgs e)
        {
            this.SuspendLayout();
        }

        private void mainFrm_ResizeEnd(object sender, EventArgs e)
        {
            this.ResumeLayout();
        }

        private void callWebDriver()
        {
            _iccesReader = new WebReader(headless: true);
            _iapmoesESWeb = new WebReader(headless: true);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Uri.IsWellFormedUriString(tbFirstWeb.Text, UriKind.Absolute))
            {
                System.Diagnostics.Process.Start(tbFirstWeb.Text);
            }
        }
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Uri.IsWellFormedUriString(tbSecondWeb.Text, UriKind.Absolute))
            {
                System.Diagnostics.Process.Start(tbSecondWeb.Text);
            }
        }
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Uri.IsWellFormedUriString(tbThirdWeb.Text, UriKind.Absolute))
            {
                System.Diagnostics.Process.Start(tbThirdWeb.Text);
            }
        }
        
        private void mainTableView_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                string link = mainTableView.CurrentCell.ToolTipText;
                if (Uri.IsWellFormedUriString(link, UriKind.Absolute))
                {
                    tbConsole.AppendText("Open Link: " + link);
                    tbConsole.AppendText(Environment.NewLine);
                    System.Diagnostics.Process.Start(link);
                }
                else 
                {
                    tbConsole.AppendText("Link is not Valid");
                    tbConsole.AppendText(Environment.NewLine);
                }
            } 
        }

        private void btnUpdateTab_Click(object sender, EventArgs e)
        {
            string xlPath = System.IO.Path.Combine(_currentDirectory, "Data", tbExcelFile.Text);
            ExcelReader xlReader = new ExcelReader(xlPath);
            if (File.Exists(xlPath))
            {
                for (int i = 0; i < cbSelectTab.Items.Count; i++)
                {
                    string currentText = cbSelectTab.GetItemText(cbSelectTab.Items[i]);
                    if ( currentText!="Setting")
                    {
                        Dictionary<string, Codes> codeReports = xlReader.getCodeDatas(currentText, tbExcelRange.Text);
                        string binPath = _currentDirectory + "//Data//" + currentText + ".bin";
                        IOData.WriteToBinaryFile(binPath, codeReports);
                        tbConsole.AppendText("Tab: " + currentText + " Has been updated.");
                        tbConsole.AppendText(Environment.NewLine);
                    }
                }
                tbConsole.AppendText("All has Done!"+ Environment.NewLine);
            }
            else
            {
                tbConsole.AppendText("Can not find Excel file.");
                tbConsole.AppendText(Environment.NewLine);
            }
            xlReader.Delete();
        }

        private void tabCompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabCompany.SelectedTab.Text == "Setting")
            {
                tabCompany.BringToFront();
            }
            else
            {
                tabCompany.SendToBack();
            }
            mainTableView.SuspendLayout();
            mainTableView.Rows.Clear();
            mainTableView.Refresh();
            string binPath = System.IO.Path.Combine(_currentDirectory, "Data", tabCompany.SelectedTab.Text + ".bin");
            if (File.Exists(binPath))
            {
                Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
                foreach (var code in codeReports)
                {
                    DataGridViewRow row = (DataGridViewRow)mainTableView.Rows[0].Clone();
                    row.Cells[0].ToolTipText = code.Value.Link;
                    row.Cells[0].Value = code.Value.Number;
                    row.Cells[1].Value = code.Value.ProductCategory;
                    row.Cells[2].Value = code.Value.Description;
                    row.Cells[3].Value = code.Value.ProductsListed;
                    row.Cells[4].Value = code.Value.LatestCode;
                    row.Cells[5].Value = code.Value.IssueDate;
                    row.Cells[6].Value = code.Value.ExpirationDate;
                    row.Cells[7].Value = code.Value.WebType;
                    row.Cells[8].ToolTipText = @"Check Web      : Not Yet" + Environment.NewLine + "Check Update : Not Yet";
                    mainTableView.Rows.Add(row);
                }
            }
            mainTableView.ResumeLayout();
            _currentPage = tabCompany.SelectedTab;
        }

        private void checkCodeUrls()
        {
            for (int i = 0; i < mainTableView.Rows.Count - 1; i++)
            {
                string webID = mainTableView.Rows[i].Cells[7].Value.ToString();
                string code = mainTableView.Rows[i].Cells[0].Value.ToString();
                string link = mainTableView.Rows[i].Cells[0].ToolTipText;
                string filename = System.IO.Path.GetFileName(link);
                string codeNo = Regex.Match(code, @"\d+").Value;
                if (!Uri.IsWellFormedUriString(mainTableView.Rows[i].Cells[0].ToolTipText, UriKind.Absolute) || !link.Contains(codeNo))
                {
                    mainTableView.Rows[i].Cells[0].Style.BackColor = btnLinkNotValid.BackColor;
                }
            }
        }

        private void mainTableView_MouseClick(object sender, MouseEventArgs e)
        {
            int currentCol = mainTableView.HitTest(e.X, e.Y).ColumnIndex;
            int currentRow = mainTableView.HitTest(e.X, e.Y).RowIndex;
            if (e.Button == MouseButtons.Right && currentCol == 0 && currentRow >=0)
            {
                mainTableView.ClearSelection();
                mainTableView.Rows[currentRow].Selected = true;
                this.tableMenu.Show(mainTableView, new System.Drawing.Point(e.X, e.Y));
            }
        }

        private async void tableMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string targetDir = tabCompany.SelectedTab.Text;
            await Task.Run(() =>
            {
                ToolStripItem item = e.ClickedItem;
                if (item.Text == "Search Code")
                {
                    int rowNo = mainTableView.SelectedRows[0].Index;
                    string code = mainTableView.Rows[rowNo].Cells[0].Value.ToString();
                    string webID = mainTableView.Rows[rowNo].Cells[7].Value.ToString();
                    checkCodeByLink(rowNo, targetDir);
                }
                if (item.Text == "Check Missing Link")
                {
                    checkCodeUrls();
                }
            });
        }

        private void listFieldMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem; 
            if (item.Text == "Add")
            {
                string value = "";
                if (InputBox("Input", "Input Field:", ref value) == DialogResult.OK)
                {
                    if (value != "" && value != null)
                    {
                        lbField.Items.Add(value);
                    }
                    else 
                    {
                        MessageBox.Show("Invalid Value!!!");
                    }
                }
            }
            if (item.Text == "Remove")
            {
                System.Windows.Forms.ListBox.SelectedObjectCollection selectedItems = new System.Windows.Forms.ListBox.SelectedObjectCollection(lbField);
                selectedItems = lbField.SelectedItems;

                if (lbField.SelectedIndex != -1)
                {
                    for (int i = selectedItems.Count - 1; i >= 0; i--)
                        lbField.Items.Remove(selectedItems[i]);
                }
            }
        }

        private void tbConsole_TextChanged(object sender, EventArgs e)
        {
            tbConsole.ScrollToCaret();
        }

        private void btnAddTab_Click(object sender, EventArgs e)
        {
            //_02_inputFrm inputFrm = new _02_inputFrm();
            //inputFrm.Show(this);
        }

        public void tableToFile()
        {
            string binPath = Path.Combine(_currentDirectory, "Data", tabCompany.SelectedTab.Text.ToString() + ".bin");
            Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
            for (int i = 0; i < mainTableView.Rows.Count - 1; i++)
            {
                string number = mainTableView.Rows[i].Cells[0].Value.ToString();
                string webType = mainTableView.Rows[i].Cells[7].Value.ToString();
                codeReports[number + "_" + webType].Link = mainTableView.Rows[i].Cells[0].ToolTipText.ToString();
            }
            IOData.WriteToBinaryFile(binPath, codeReports);
        }

        private void checkCodeByLink(int row, string targetDir)
        {
            string output = "";
            string link = mainTableView.Rows[row].Cells[0].ToolTipText.ToString();
            string webID = mainTableView.Rows[row].Cells[7].Value.ToString();
            string code = mainTableView.Rows[row].Cells[0].Value.ToString();
            if (webID == "ICC-ES ESR")
            {
                output = _iccesReader.trackICCESReport(tbThirdWeb.Text, code, targetDir);
            }
            else
            {
                output = _iapmoesESWeb.trackIAPMOESReport(tbSecondWeb.Text, code, targetDir);
            }


            if (output == "No Reports found")
            {
                mainTableView.Rows[row].Cells[0].Style.BackColor = btnLinkNotFound.BackColor;
            }
            if (output.Contains("Cancelled"))
            {
                mainTableView.Rows[row].Cells[0].Style.BackColor = btnLinkCancelled.BackColor;
            }
            else if (output != link && output != "No Reports found")
            {
                mainTableView.Rows[row].Cells[0].Style.BackColor = btnLinkUpdated.BackColor;
            }
            mainTableView.Rows[row].Cells[0].ToolTipText = output;
            mainTableView.Rows[row].Cells[8].Value = true;
            downloadFileAsync(mainTableView.Rows[row].Cells[0].ToolTipText, _currentDirectory + "//Data//" + targetDir + "//" + code + "_" + webID + ".pdf", row);
            mainTableView.Rows[row].Cells[8].ToolTipText = updateStatus(mainTableView.Rows[row].Cells[8].ToolTipText, 1);
        }

        private async void btnCheckCodes_Click(object sender, EventArgs e)
        {
            _webThread.Join();
            if (tabCompany.SelectedTab.Text != "Setting")
            {
                tbConsole.AppendText("Start Running...");
                tbConsole.AppendText(Environment.NewLine);
                blockItems();
                await SearchReport();
                tableToFile();
                unBlockItems();
                tbConsole.AppendText("Done.");
                tbConsole.AppendText(Environment.NewLine);
            }
        }

        private async void btnUpdateDate_Click(object sender, EventArgs e)
        {
            blockItems();
            await checkDate();
            unBlockItems();
        }

        private void btnExcelExport_Click(object sender, EventArgs e)
        {
            string targetPath = Path.Combine(_currentDirectory, "Data", "ExportData.xlsx");
            string binPath = Path.Combine(_currentDirectory, "Data", tabCompany.SelectedTab.Text.ToString() + ".bin");
            ExcelReader xlReader = new ExcelReader(targetPath, FileAccess.ReadWrite);
            string currentTime = DateTime.Now.ToString("MM/dd/yyyy h:mm tt");
            xlReader.selectSheet(tabCompany.SelectedTab.Text.ToString());
            int maxRow = xlReader.maxRow();
            xlReader.writeCellData(maxRow, 1, "Last Recorded:");
            xlReader.writeCellData(maxRow, 2, currentTime);
            xlReader.writeCellColor(maxRow, 1, System.Drawing.Color.Aqua);
            xlReader.writeCellColor(maxRow, 2, System.Drawing.Color.Aqua);

            //Dictionary<string, Codes> codeData = new Dictionary<string, Codes>();
            Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);

            for (int i = 0; i < mainTableView.Rows.Count - 1; i++)
            {
                for (int j = 0; j < 7; j++)
                {
                    string value = mainTableView.Rows[i].Cells[j].Value?.ToString() ?? "";
                    System.Drawing.Color cellColor = mainTableView.Rows[i].Cells[j].Style.BackColor;
                    xlReader.writeCellData(maxRow + i + 1, j + 1, value);
                    if (cellColor != System.Drawing.Color.Empty)
                    {
                        xlReader.writeCellColor(maxRow + i + 1, j + 1, cellColor);
                    }
                    if (mainTableView.Rows[i].Cells[j].ToolTipText != "")
                    {
                        xlReader.writeHyperLink(maxRow + i + 1, 1, mainTableView.Rows[i].Cells[j].ToolTipText);
                    }
                }

                string number = mainTableView.Rows[i].Cells[0].Value.ToString();
                string webType = mainTableView.Rows[i].Cells[7].Value.ToString();
                codeReports[number + "_" + webType].Number = mainTableView.Rows[i].Cells[0].Value.ToString();
                codeReports[number + "_" + webType].Link = mainTableView.Rows[i].Cells[0].ToolTipText.ToString();
                codeReports[number + "_" + webType].ProductCategory = mainTableView.Rows[i].Cells[1].Value.ToString();
                codeReports[number + "_" + webType].Description = mainTableView.Rows[i].Cells[2].Value.ToString();
                codeReports[number + "_" + webType].ProductsListed = mainTableView.Rows[i].Cells[3].Value.ToString();
                codeReports[number + "_" + webType].LatestCode = mainTableView.Rows[i].Cells[4].Value.ToString();
                codeReports[number + "_" + webType].IssueDate = mainTableView.Rows[i].Cells[5].Value.ToString();
                codeReports[number + "_" + webType].ExpirationDate = mainTableView.Rows[i].Cells[6].Value.ToString();
                codeReports[number + "_" + webType].WebType = mainTableView.Rows[i].Cells[7].Value.ToString();
                
            }
            IOData.WriteToBinaryFile(binPath, codeReports);
            tbConsole.AppendText("Sheet: " + tabCompany.SelectedTab.Text.ToString() + " Was Exported.");
            tbConsole.AppendText(Environment.NewLine);
            xlReader.saveWorkBook();
            xlReader.Delete();
        }

        public void downloadFileAsync(string Url, string fileName, int row)
        {
            if (Uri.IsWellFormedUriString(Url, UriKind.Absolute))
            {
                WebClient webClient = new WebClient();
                webClient.DownloadFileCompleted += webClientDownloadFileCompleted;
                webClient.DownloadProgressChanged += webClientDownloadProgressChanged;
                webClient.QueryString.Add("row", row.ToString());
                webClient.DownloadFileAsync(new Uri(Url), fileName);
            }   
        }

        private void webClientDownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            string row = ((System.Net.WebClient)(sender)).QueryString["row"];
            int rowNo = Int32.Parse(row);
            mainTableView.Rows[rowNo].Cells[9].Value = 100;
        }

        private void webClientDownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            string row = ((System.Net.WebClient)(sender)).QueryString["row"];
            int rowNo = Int32.Parse(row);
            int percent = Convert.ToInt32((e.BytesReceived / e.TotalBytesToReceive) * 100);
            mainTableView.Rows[rowNo].Cells[9].Value = percent;
        }

        #region Task
        public async Task SearchReport()
        {
            string targetDir = tabCompany.SelectedTab.Text;
            for (int i = 0; i < mainTableView.Rows.Count - 1; i++)
            {
                int percent = i * 101 / (mainTableView.Rows.Count - 1);
                pBarState.Value = percent;
                pBarState.CustomText = percent.ToString();
                string webID = mainTableView.Rows[i].Cells[7].Value.ToString();
                string code = mainTableView.Rows[i].Cells[0].Value.ToString();
                if (rBtnUseWeb.Checked)
                {
                    await Task.Run(() =>
                    {
                        checkCodeByLink(i, targetDir);
                    });
                }
                else
                {
                    await Task.Run(() =>
                    {
                        if (webID == "ICC-ES ESR")
                        {
                            string targetLink = "https://cdn-v2.icc-es.org/wp-content/uploads/report-directory/" + code + ".pdf";
                            if (_iccesReader.RemoteFileExists(targetLink))
                            {
                                downloadFileAsync(targetLink, _currentDirectory + "//Data//" + targetDir + "//" + code + "_" + webID + ".pdf", i);
                            }
                            else
                            {
                                mainTableView.Rows[i].Cells[0].Style.BackColor = btnLinkNotValid.BackColor;
                            }
                            mainTableView.Rows[i].Cells[8].Value = true;
                        }
                        if (webID == "IAPMO UES ER")
                        {
                            string targetLink = mainTableView.Rows[i].Cells[0].ToolTipText.ToString();
                            if (_iapmoesESWeb.RemoteFileExists(targetLink))
                            {
                                downloadFileAsync(targetLink, _currentDirectory + "//Data//" + targetDir + "//" + code + "_" + webID + ".pdf", i);
                            }
                            else
                            {
                                mainTableView.Rows[i].Cells[0].Style.BackColor = btnLinkNotFound.BackColor;
                            }
                            mainTableView.Rows[i].Cells[8].Value = true;
                        }
                    });
                }
                pBarState.Value = 100;
                pBarState.CustomText = "100%";
            }
        }

        public async Task checkDate()
        {
            string tabName = tabCompany.SelectedTab.Text;
            for (int i = 0; i < mainTableView.Rows.Count - 1; i++)
            {
                int percent = i * 101 / (mainTableView.Rows.Count - 1);
                pBarState.Value = percent;
                pBarState.CustomText = percent.ToString();
                string code = mainTableView.Rows[i].Cells[0].Value.ToString();
                string webID = mainTableView.Rows[i].Cells[7].Value.ToString();
                string file = Path.Combine(_currentDirectory, "Data", tabName, code + "_" + webID + ".pdf");
                await Task.Run(() =>
                {
                    if (File.Exists(file))
                    {
                        try
                        {
                            PDFReader pDFReader = new PDFReader(file);
                            List<string> items = new List<string>();
                            if (mainTableView.Rows[i].Cells[7].Value.ToString() == "ICC-ES ESR")
                            {
                                items = pDFReader.iccesSearch();
                            }
                            else
                            {
                                items = pDFReader.iapmoSearch();
                            }

                            if (items[0] != "" && items[0].Trim().ToLower() != mainTableView.Rows[i].Cells[4].Value.ToString().Trim().ToLower())
                            {
                                mainTableView.Rows[i].Cells[4].Value = items[0];
                                mainTableView.Rows[i].Cells[4].Style.BackColor = btnLinkUpdated.BackColor;
                            }
                            if (items[1] != "" && items[1].Trim().ToLower() != mainTableView.Rows[i].Cells[5].Value.ToString().Trim().ToLower())
                            {
                                mainTableView.Rows[i].Cells[5].Value = items[1];
                                mainTableView.Rows[i].Cells[5].Style.BackColor = btnLinkUpdated.BackColor;
                            }
                            if (items[2] != "" && items[2].Trim().ToLower() != mainTableView.Rows[i].Cells[6].Value.ToString().Trim().ToLower())
                            {
                                mainTableView.Rows[i].Cells[6].Value = items[2];
                                mainTableView.Rows[i].Cells[6].Style.BackColor = btnLinkUpdated.BackColor;
                            }
                            pDFReader.Close();
                        }
                        catch (Exception ex)
                        {
                            //mainTableView.Rows[i].Cells[8].ToolTipText = "File:" + file + " is Error";
                            mainTableView.Rows[i].Cells[8].Style.BackColor = btnLinkNotValid.BackColor;
                        }
                    }
                    mainTableView.Rows[i].Cells[8].ToolTipText = updateStatus(mainTableView.Rows[i].Cells[8].ToolTipText, 2);
                });
                pBarState.Value = 100;
                pBarState.CustomText = "100%";
            }
        }
        #endregion

        private void disableTabs()
        {
            foreach (TabPage page in tabCompany.TabPages)
            {
                ((System.Windows.Forms.Control)page).Enabled = tabCompany.SelectedTab == page;
            }
        }
        
        private void enableTabs()
        {
            foreach (TabPage tbp in tabCompany.TabPages)
            {
                tbp.Show();
            }
        }

        private void blockItems()
        {
            btnCheckCode.Enabled = false;
            btnUpdateDate.Enabled = false;
            btnExcelExport.Enabled = false;
            tabCompany.Enabled = false;
        }
        
        private void unBlockItems()
        {
            btnCheckCode.Enabled = true;
            btnUpdateDate.Enabled = true;
            btnExcelExport.Enabled = true;
            tabCompany.Enabled = true;
        }

        private string updateStatus(string oldText, int index)
        {
            string newText = "";
            if (index == 1)
            {
                string pattern = @"(Check Web\s+:\s)(.*)";
                newText = Regex.Replace(oldText, pattern, "$1Done.");
            }
            else if (index == 2)
            {
                string pattern = @"(Check Update\s+:\s)(.*)";
                newText = Regex.Replace(oldText, pattern, "$1Done.");
            }
            else
            { newText = ""; }
            return newText;
        }

        #region button color
        private void btnLinkUpdated_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                btnLinkUpdated.BackColor = colorDialog1.Color;
            }
        }

        private void btnLinkNotFound_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                btnLinkNotFound.BackColor = colorDialog1.Color;
            }
        }
        #endregion

        private void btnLinkNotValid_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                btnLinkNotValid.BackColor = colorDialog1.Color;
            }
        }

        private void btnLinkCancelled_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                btnLinkCancelled.BackColor = colorDialog1.Color;
            }
        }

        private void btnLinkCorrect_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                btnLinkCorrect.BackColor = colorDialog1.Color;
            }
        }
    }
}
