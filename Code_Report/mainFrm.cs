using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq.Expressions;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Code_Report.Properties;
using Microsoft.CSharp.RuntimeBinder;

namespace Code_Report;

public class mainFrm : Form
{
	[CompilerGenerated]
	private static class _003C_003Eo__8
	{
		public static CallSite<Func<CallSite, object, DialogResult, object>> _003C_003Ep__0;

		public static CallSite<Func<CallSite, object, bool>> _003C_003Ep__1;

		public static CallSite<Func<CallSite, object, DialogResult, object>> _003C_003Ep__2;

		public static CallSite<Func<CallSite, object, bool>> _003C_003Ep__3;
	}

	public string _currentDirectory = Directory.GetCurrentDirectory();

	public string _currentTab;

	private WebReader _iccesReader;

	private WebReader _iapmoesESWeb;

	private Dictionary<string, int> _downloadObject;

	private IContainer components = null;

	private TextBox tbFirstWeb;

	private Panel panel1;

	private LinkLabel linkLabelFirstWeb;

	private Panel panel2;

	private LinkLabel linkLabelSecondWeb;

	private TextBox tbSecondWeb;

	private TextBox tbConsole;

	private TabPage tabPage10;

	private Button btnUpdateTab;

	private ComboBox cbSelectTab;

	private Label labelSelectTab;

	private TabPage tabPage11;

	private TabPage tabPage12;

	private TabPage tabPage9;

	private TabPage tabPage8;

	private TabPage tabPage7;

	private TabPage tabPage6;

	private TabPage tabPage5;

	private TabPage tabPage4;

	private TabPage tabPage3;

	private TabPage tabPage2;

	private TabPage tabPage1;

	private DataGridView mainTableView;

	private TabControl tabCompany;

	private GroupBox gbWebOptions;

	private Panel panel4;

	private RadioButton rBtnUseLink;

	private RadioButton rBtnUseWeb;

	private GroupBox gbTools;

	private Panel panel5;

	private Label labelExcelFile;

	private TextBox tbExcelFile;

	private Label labelExcelRange;

	private TextBox tbExcelRange;

	private Label labelCheckBy;

	private Button btnCheckLink;

	private ImageList imageList1;

	private ContextMenuStrip tableMenu;

	private ToolStripMenuItem checkLink;

	private ListBox lbField;

	private Label labelField;

	private ContextMenuStrip listFieldMenu;

	private ToolStripMenuItem toolStripMenuItem1;

	private ToolStripMenuItem toolStripMenuItem2;

	private ToolStripMenuItem checkValidLink;

	private Button btnAddTab;

	private TextBox tbThirdWeb;

	private LinkLabel linkLabelThirdWeb;

	private Panel panel3;

	private Button btnCheckCode;

	private Button btnLinkUpdated;

	private Label labelLinkUpdated;

	private Label labelLinkNotFound;

	private Button btnLinkNotFound;

	private Label labelLinkNotValid;

	private Button btnLinkNotValid;

	private Label labelLinkCancelled;

	private Button btnLinkCancelled;

	private Label labelLinkCorrect;

	private Button btnLinkCorrect;

	private CustomProgressBar pBarState;

	private Button btnUpdateDate;

	private Button btnExcelExport;

	private DataGridViewProgressColumn dataGridViewProgressColumn1;

	private DataGridViewLinkColumn Column1;

	private DataGridViewTextBoxColumn Column2;

	private DataGridViewTextBoxColumn Column3;

	private DataGridViewTextBoxColumn Column4;

	private DataGridViewTextBoxColumn Column5;

	private DataGridViewTextBoxColumn Column6;

	private DataGridViewTextBoxColumn Column7;

	private DataGridViewTextBoxColumn Type;

	private DataGridViewCheckBoxColumn Status;

	private DataGridViewProgressColumn Column8;

	public static DialogResult InputBox(string title, string promptText, ref string value)
	{
		Form form = new Form();
		Label label = new Label();
		TextBox textBox = new TextBox();
		Button buttonOk = new Button();
		Button buttonCancel = new Button();
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
		textBox.Anchor |= AnchorStyles.Right;
		buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
		form.ClientSize = new Size(396, 107);
		form.Controls.AddRange(new Control[4] { label, textBox, buttonOk, buttonCancel });
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
			Directory.CreateDirectory(targetDir);
		}
		string binPath = Directory.GetCurrentDirectory() + "//Data//SST ESRs & ERs.bin";
		if (File.Exists(binPath))
		{
			Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
			foreach (KeyValuePair<string, Codes> code in codeReports)
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
				mainTableView.Rows.Add(row);
			}
		}
		mainTableView.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
		_iccesReader = new WebReader(headless: true);
		_iapmoesESWeb = new WebReader(headless: true);
	}

	private void mainFrm_FormClosing(object sender, FormClosingEventArgs e)
	{
		if (e.CloseReason != CloseReason.UserClosing)
		{
			return;
		}
		object result = MessageBox.Show("Do You Want To Close Program?", "Code Report", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
		if (_003C_003Eo__8._003C_003Ep__1 == null)
		{
			_003C_003Eo__8._003C_003Ep__1 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation((CSharpBinderFlags)0, ExpressionType.IsTrue, typeof(mainFrm), (IEnumerable<CSharpArgumentInfo>)(object)new CSharpArgumentInfo[1] { CSharpArgumentInfo.Create((CSharpArgumentInfoFlags)0, (string)null) }));
		}
		Func<CallSite, object, bool> target = _003C_003Eo__8._003C_003Ep__1.Target;
		CallSite<Func<CallSite, object, bool>> _003C_003Ep__ = _003C_003Eo__8._003C_003Ep__1;
		if (_003C_003Eo__8._003C_003Ep__0 == null)
		{
			_003C_003Eo__8._003C_003Ep__0 = CallSite<Func<CallSite, object, DialogResult, object>>.Create(Binder.BinaryOperation((CSharpBinderFlags)0, ExpressionType.Equal, typeof(mainFrm), (IEnumerable<CSharpArgumentInfo>)(object)new CSharpArgumentInfo[2]
			{
				CSharpArgumentInfo.Create((CSharpArgumentInfoFlags)0, (string)null),
				CSharpArgumentInfo.Create((CSharpArgumentInfoFlags)3, (string)null)
			}));
		}
		if (target(_003C_003Ep__, _003C_003Eo__8._003C_003Ep__0.Target(_003C_003Eo__8._003C_003Ep__0, result, DialogResult.Yes)))
		{
			if (_iccesReader != null)
			{
				_iccesReader.closeDriver();
			}
			if (_iapmoesESWeb != null)
			{
				_iapmoesESWeb.closeDriver();
			}
			Application.Exit();
		}
		if (_003C_003Eo__8._003C_003Ep__3 == null)
		{
			_003C_003Eo__8._003C_003Ep__3 = CallSite<Func<CallSite, object, bool>>.Create(Binder.UnaryOperation((CSharpBinderFlags)0, ExpressionType.IsTrue, typeof(mainFrm), (IEnumerable<CSharpArgumentInfo>)(object)new CSharpArgumentInfo[1] { CSharpArgumentInfo.Create((CSharpArgumentInfoFlags)0, (string)null) }));
		}
		Func<CallSite, object, bool> target2 = _003C_003Eo__8._003C_003Ep__3.Target;
		CallSite<Func<CallSite, object, bool>> _003C_003Ep__2 = _003C_003Eo__8._003C_003Ep__3;
		if (_003C_003Eo__8._003C_003Ep__2 == null)
		{
			_003C_003Eo__8._003C_003Ep__2 = CallSite<Func<CallSite, object, DialogResult, object>>.Create(Binder.BinaryOperation((CSharpBinderFlags)0, ExpressionType.Equal, typeof(mainFrm), (IEnumerable<CSharpArgumentInfo>)(object)new CSharpArgumentInfo[2]
			{
				CSharpArgumentInfo.Create((CSharpArgumentInfoFlags)0, (string)null),
				CSharpArgumentInfo.Create((CSharpArgumentInfoFlags)3, (string)null)
			}));
		}
		if (target2(_003C_003Ep__2, _003C_003Eo__8._003C_003Ep__2.Target(_003C_003Eo__8._003C_003Ep__2, result, DialogResult.No)))
		{
			e.Cancel = true;
		}
	}

	private void mainFrm_ResizeBegin(object sender, EventArgs e)
	{
		SuspendLayout();
	}

	private void mainFrm_ResizeEnd(object sender, EventArgs e)
	{
		ResumeLayout();
	}

	private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
	{
		if (Uri.IsWellFormedUriString(tbFirstWeb.Text, UriKind.Absolute))
		{
			Process.Start(tbFirstWeb.Text);
		}
	}

	private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
	{
		if (Uri.IsWellFormedUriString(tbSecondWeb.Text, UriKind.Absolute))
		{
			Process.Start(tbSecondWeb.Text);
		}
	}

	private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
	{
		if (Uri.IsWellFormedUriString(tbThirdWeb.Text, UriKind.Absolute))
		{
			Process.Start(tbThirdWeb.Text);
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
				Process.Start(link);
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
		ExcelReader xlReader = new ExcelReader();
		string xlPath = Path.Combine(_currentDirectory, "Data", tbExcelFile.Text);
		if (File.Exists(xlPath))
		{
			xlReader.loadWorkBook(xlPath);
			for (int i = 0; i < cbSelectTab.Items.Count; i++)
			{
				string currentText = cbSelectTab.GetItemText(cbSelectTab.Items[i]);
				if (currentText != "Setting")
				{
					Dictionary<string, Codes> codeReports = xlReader.getTableByRange(currentText, tbExcelRange.Text);
					string binPath = _currentDirectory + "//Data//" + currentText + ".bin";
					IOData.WriteToBinaryFile(binPath, codeReports);
					tbConsole.AppendText("Tab: " + currentText + " Has been updated.");
					tbConsole.AppendText(Environment.NewLine);
				}
			}
			tbConsole.AppendText("All has Done!");
		}
		else
		{
			tbConsole.AppendText("Can not find Excel file.");
			tbConsole.AppendText(Environment.NewLine);
		}
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
		string binPath = Path.Combine(_currentDirectory, "Data", tabCompany.SelectedTab.Text + ".bin");
		if (File.Exists(binPath))
		{
			Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
			foreach (KeyValuePair<string, Codes> code in codeReports)
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
				mainTableView.Rows.Add(row);
			}
		}
		mainTableView.ResumeLayout();
	}

	private void btnCheckLink_Click()
	{
		for (int i = 0; i < mainTableView.Rows.Count - 1; i++)
		{
			string webID = mainTableView.Rows[i].Cells[7].Value.ToString();
			if (webID != "LADBS RR")
			{
				string code = mainTableView.Rows[i].Cells[0].Value.ToString();
				string link = mainTableView.Rows[i].Cells[0].ToolTipText;
				string filename = Path.GetFileName(link);
				string codeNo = Regex.Match(code, "\\d+").Value;
				if (!Uri.IsWellFormedUriString(mainTableView.Rows[i].Cells[0].ToolTipText, UriKind.Absolute) || !link.Contains(codeNo))
				{
					mainTableView.Rows[i].Cells[0].Style.BackColor = Color.Red;
				}
			}
		}
	}

	private void mainTableView_MouseClick(object sender, MouseEventArgs e)
	{
		int currentCol = mainTableView.HitTest(e.X, e.Y).ColumnIndex;
		int currentRow = mainTableView.HitTest(e.X, e.Y).RowIndex;
		if (e.Button == MouseButtons.Right && currentCol == 0 && currentRow >= 0)
		{
			mainTableView.ClearSelection();
			mainTableView.Rows[currentRow].Selected = true;
			tableMenu.Show(mainTableView, new Point(e.X, e.Y));
		}
	}

	private async void tableMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
	{
		string targetDir = tabCompany.SelectedTab.Text;
		await Task.Run(delegate
		{
			ToolStripItem clickedItem = e.ClickedItem;
			if (clickedItem.Text == "Search Code")
			{
				int index = mainTableView.SelectedRows[0].Index;
				string reportNumber = mainTableView.Rows[index].Cells[0].Value.ToString();
				string text = mainTableView.Rows[index].Cells[7].Value.ToString();
				if (text == "IAPMO UES ER")
				{
					_iccesReader.trackIAPMOESReport(tbSecondWeb.Text, reportNumber, targetDir);
				}
				if (text == "ICC-ES ESR")
				{
					string text2 = _iccesReader.trackICCESReport(tbThirdWeb.Text, reportNumber, targetDir);
					if (text2 == "No Reports found")
					{
						mainTableView.Rows[index].Cells[0].Style.BackColor = Color.Orange;
					}
					else if (text2 != mainTableView.Rows[index].Cells[0].ToolTipText && text2 != "No Reports found")
					{
						mainTableView.Rows[index].Cells[0].ToolTipText = text2;
						mainTableView.Rows[index].Cells[0].Style.BackColor = Color.Green;
					}
					mainTableView.Rows[index].Cells[8].Value = true;
				}
			}
			if (clickedItem.Text == "Check Missing Link")
			{
				btnCheckLink_Click();
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
		if (!(item.Text == "Remove"))
		{
			return;
		}
		ListBox.SelectedObjectCollection selectedItems = new ListBox.SelectedObjectCollection(lbField);
		selectedItems = lbField.SelectedItems;
		if (lbField.SelectedIndex != -1)
		{
			for (int i = selectedItems.Count - 1; i >= 0; i--)
			{
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
	}

	public async Task SearchReport()
	{
		string targetDir = tabCompany.SelectedTab.Text;
		for (int i = 0; i < mainTableView.Rows.Count - 1; i++)
		{
			int percent = i * 101 / (mainTableView.Rows.Count - 1);
			pBarState.Value = percent;
			pBarState.CustomText = percent.ToString();
			string output = "";
			string webID = mainTableView.Rows[i].Cells[7].Value.ToString();
			string code = mainTableView.Rows[i].Cells[0].Value.ToString();
			if (rBtnUseWeb.Checked)
			{
				await Task.Run(delegate
				{
					string text = mainTableView.Rows[i].Cells[0].ToolTipText.ToString();
					if (webID == "ICC-ES ESR")
					{
						output = _iccesReader.trackICCESReport(tbThirdWeb.Text, code, targetDir);
						if (output == "No Reports found")
						{
							mainTableView.Rows[i].Cells[0].Style.BackColor = Color.Orange;
						}
						else if (output != text && output != "No Reports found")
						{
							mainTableView.Rows[i].Cells[0].ToolTipText = output;
							mainTableView.Rows[i].Cells[0].Style.BackColor = Color.LightGreen;
						}
						mainTableView.Rows[i].Cells[8].Value = true;
						downloadFileAsync(mainTableView.Rows[i].Cells[0].ToolTipText, _currentDirectory + "//Data//" + targetDir + "//" + code + ".pdf", i);
					}
					if (webID == "IAPMO UES ER")
					{
						output = _iapmoesESWeb.trackIAPMOESReport(tbSecondWeb.Text, code, targetDir);
						if (output == "No Reports found")
						{
							mainTableView.Rows[i].Cells[0].Style.BackColor = Color.Orange;
						}
						if (output.Contains("Cancelled "))
						{
							mainTableView.Rows[i].Cells[0].Style.BackColor = Color.Gray;
							mainTableView.Rows[i].Cells[0].ToolTipText = output;
						}
						else if (output != text && output != "No Reports found")
						{
							mainTableView.Rows[i].Cells[0].ToolTipText = output;
							mainTableView.Rows[i].Cells[0].Style.BackColor = Color.LightGreen;
						}
						mainTableView.Rows[i].Cells[8].Value = true;
						downloadFileAsync(mainTableView.Rows[i].Cells[0].ToolTipText, _currentDirectory + "//Data//" + targetDir + "//" + code + ".pdf", i);
					}
				});
			}
			else
			{
				await Task.Run(delegate
				{
					if (webID == "ICC-ES ESR")
					{
						string url = "https://cdn-v2.icc-es.org/wp-content/uploads/report-directory/" + code + ".pdf";
						if (_iccesReader.RemoteFileExists(url))
						{
							downloadFileAsync(url, _currentDirectory + "//Data//" + targetDir + "//" + code + ".pdf", i);
						}
						else
						{
							mainTableView.Rows[i].Cells[0].Style.BackColor = Color.Red;
						}
						mainTableView.Rows[i].Cells[8].Value = true;
					}
					if (webID == "IAPMO UES ER")
					{
						string url2 = mainTableView.Rows[i].Cells[0].ToolTipText.ToString();
						if (_iapmoesESWeb.RemoteFileExists(url2))
						{
							downloadFileAsync(url2, _currentDirectory + "//Data//" + targetDir + "//" + code + ".pdf", i);
						}
						else
						{
							mainTableView.Rows[i].Cells[0].Style.BackColor = Color.Red;
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
			mainTableView.Rows[i].Cells[9].Value = percent;
			pBarState.Value = percent;
			pBarState.CustomText = percent.ToString();
			string file = Path.Combine(path4: mainTableView.Rows[i].Cells[0].Value.ToString() + ".pdf", path1: _currentDirectory, path2: "Data", path3: tabName);
			await Task.Run(delegate
			{
				if (File.Exists(file))
				{
					PDFReader pDFReader = new PDFReader(file);
					List<string> list = new List<string>();
					if (mainTableView.Rows[i].Cells[7].Value.ToString() == "ICC-ES ESR")
					{
						list = pDFReader.iccesSearch();
						if (list[0] != "" && list[0] != mainTableView.Rows[i].Cells[4].Value.ToString())
						{
							mainTableView.Rows[i].Cells[4].Value = list[0];
							mainTableView.Rows[i].Cells[4].Style.BackColor = Color.LightGreen;
						}
						if (list[1] != "" && list[1] != mainTableView.Rows[i].Cells[5].Value.ToString())
						{
							mainTableView.Rows[i].Cells[5].Value = list[1];
							mainTableView.Rows[i].Cells[5].Style.BackColor = Color.LightGreen;
						}
						if (list[2] != "" && list[2] != mainTableView.Rows[i].Cells[6].Value.ToString())
						{
							mainTableView.Rows[i].Cells[6].Value = list[2];
							mainTableView.Rows[i].Cells[6].Style.BackColor = Color.LightGreen;
						}
					}
					if (mainTableView.Rows[i].Cells[7].Value.ToString() == "IAPMO UES ER")
					{
						list = pDFReader.iapmoSearch();
						if (list[0] != "" && list[0] != mainTableView.Rows[i].Cells[4].Value.ToString())
						{
							mainTableView.Rows[i].Cells[4].Value = list[0];
							mainTableView.Rows[i].Cells[4].Style.BackColor = Color.LightGreen;
						}
						if (list[1] != "" && list[1] != mainTableView.Rows[i].Cells[5].Value.ToString())
						{
							mainTableView.Rows[i].Cells[5].Value = list[1];
							mainTableView.Rows[i].Cells[5].Style.BackColor = Color.LightGreen;
						}
						if (list[2] != "" && list[2] != mainTableView.Rows[i].Cells[6].Value.ToString())
						{
							mainTableView.Rows[i].Cells[6].Value = list[2];
							mainTableView.Rows[i].Cells[6].Style.BackColor = Color.LightGreen;
						}
					}
					pDFReader.Close();
				}
			});
			pBarState.Value = 100;
			pBarState.CustomText = "100%";
		}
	}

	private async void btnCheckCodes_Click(object sender, EventArgs e)
	{
		tbConsole.AppendText("Start Running...");
		tbConsole.AppendText(Environment.NewLine);
		await SearchReport();
		tbConsole.AppendText("Done!!!");
		tbConsole.AppendText(Environment.NewLine);
	}

	private async void btnUpdateDate_Click(object sender, EventArgs e)
	{
		await checkDate();
	}

	public void downloadFileAsync(string Url, string fileName, int row)
	{
		WebClient webClient = new WebClient();
		webClient.DownloadFileCompleted += webClientDownloadFileCompleted;
		webClient.DownloadProgressChanged += webClientDownloadProgressChanged;
		webClient.QueryString.Add("row", row.ToString());
		webClient.DownloadFileAsync(new Uri(Url), fileName);
	}

	private void webClientDownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
	{
		string row = ((WebClient)sender).QueryString["row"];
		int rowNo = int.Parse(row);
		mainTableView.Rows[rowNo].Cells[9].Value = 100;
	}

	private void webClientDownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
	{
		string row = ((WebClient)sender).QueryString["row"];
		int rowNo = int.Parse(row);
		mainTableView.Rows[rowNo].Cells[9].Value = e.BytesReceived / e.TotalBytesToReceive * 100;
		Console.WriteLine(rowNo);
	}

	protected override void Dispose(bool disposing)
	{
		if (disposing && components != null)
		{
			components.Dispose();
		}
		base.Dispose(disposing);
	}

	private void InitializeComponent()
	{
		this.components = new System.ComponentModel.Container();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle12 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle21 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle22 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle13 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle14 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle15 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle16 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle17 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle18 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle19 = new System.Windows.Forms.DataGridViewCellStyle();
		System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle20 = new System.Windows.Forms.DataGridViewCellStyle();
		System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Code_Report.mainFrm));
		this.tbFirstWeb = new System.Windows.Forms.TextBox();
		this.panel1 = new System.Windows.Forms.Panel();
		this.linkLabelFirstWeb = new System.Windows.Forms.LinkLabel();
		this.panel2 = new System.Windows.Forms.Panel();
		this.linkLabelSecondWeb = new System.Windows.Forms.LinkLabel();
		this.tbSecondWeb = new System.Windows.Forms.TextBox();
		this.tbConsole = new System.Windows.Forms.TextBox();
		this.tabPage10 = new System.Windows.Forms.TabPage();
		this.gbTools = new System.Windows.Forms.GroupBox();
		this.panel5 = new System.Windows.Forms.Panel();
		this.btnAddTab = new System.Windows.Forms.Button();
		this.labelField = new System.Windows.Forms.Label();
		this.lbField = new System.Windows.Forms.ListBox();
		this.listFieldMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
		this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
		this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
		this.btnCheckLink = new System.Windows.Forms.Button();
		this.tbExcelRange = new System.Windows.Forms.TextBox();
		this.labelExcelRange = new System.Windows.Forms.Label();
		this.tbExcelFile = new System.Windows.Forms.TextBox();
		this.labelExcelFile = new System.Windows.Forms.Label();
		this.btnUpdateTab = new System.Windows.Forms.Button();
		this.cbSelectTab = new System.Windows.Forms.ComboBox();
		this.labelSelectTab = new System.Windows.Forms.Label();
		this.gbWebOptions = new System.Windows.Forms.GroupBox();
		this.panel4 = new System.Windows.Forms.Panel();
		this.labelCheckBy = new System.Windows.Forms.Label();
		this.rBtnUseLink = new System.Windows.Forms.RadioButton();
		this.rBtnUseWeb = new System.Windows.Forms.RadioButton();
		this.tabPage11 = new System.Windows.Forms.TabPage();
		this.tabPage12 = new System.Windows.Forms.TabPage();
		this.tabPage9 = new System.Windows.Forms.TabPage();
		this.tabPage8 = new System.Windows.Forms.TabPage();
		this.tabPage7 = new System.Windows.Forms.TabPage();
		this.tabPage6 = new System.Windows.Forms.TabPage();
		this.tabPage5 = new System.Windows.Forms.TabPage();
		this.tabPage4 = new System.Windows.Forms.TabPage();
		this.tabPage3 = new System.Windows.Forms.TabPage();
		this.tabPage2 = new System.Windows.Forms.TabPage();
		this.tabPage1 = new System.Windows.Forms.TabPage();
		this.mainTableView = new System.Windows.Forms.DataGridView();
		this.Column1 = new System.Windows.Forms.DataGridViewLinkColumn();
		this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Column4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Column5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Column6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Column7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Type = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.Status = new System.Windows.Forms.DataGridViewCheckBoxColumn();
		this.Column8 = new Code_Report.DataGridViewProgressColumn();
		this.tableMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
		this.checkLink = new System.Windows.Forms.ToolStripMenuItem();
		this.checkValidLink = new System.Windows.Forms.ToolStripMenuItem();
		this.tabCompany = new System.Windows.Forms.TabControl();
		this.imageList1 = new System.Windows.Forms.ImageList(this.components);
		this.tbThirdWeb = new System.Windows.Forms.TextBox();
		this.linkLabelThirdWeb = new System.Windows.Forms.LinkLabel();
		this.panel3 = new System.Windows.Forms.Panel();
		this.btnCheckCode = new System.Windows.Forms.Button();
		this.btnLinkUpdated = new System.Windows.Forms.Button();
		this.labelLinkUpdated = new System.Windows.Forms.Label();
		this.labelLinkNotFound = new System.Windows.Forms.Label();
		this.btnLinkNotFound = new System.Windows.Forms.Button();
		this.labelLinkNotValid = new System.Windows.Forms.Label();
		this.btnLinkNotValid = new System.Windows.Forms.Button();
		this.labelLinkCancelled = new System.Windows.Forms.Label();
		this.btnLinkCancelled = new System.Windows.Forms.Button();
		this.labelLinkCorrect = new System.Windows.Forms.Label();
		this.btnLinkCorrect = new System.Windows.Forms.Button();
		this.btnUpdateDate = new System.Windows.Forms.Button();
		this.btnExcelExport = new System.Windows.Forms.Button();
		this.dataGridViewProgressColumn1 = new Code_Report.DataGridViewProgressColumn();
		this.pBarState = new Code_Report.CustomProgressBar();
		this.panel1.SuspendLayout();
		this.panel2.SuspendLayout();
		this.tabPage10.SuspendLayout();
		this.gbTools.SuspendLayout();
		this.panel5.SuspendLayout();
		this.listFieldMenu.SuspendLayout();
		this.gbWebOptions.SuspendLayout();
		this.panel4.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)this.mainTableView).BeginInit();
		this.tableMenu.SuspendLayout();
		this.tabCompany.SuspendLayout();
		this.panel3.SuspendLayout();
		base.SuspendLayout();
		this.tbFirstWeb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.tbFirstWeb.Location = new System.Drawing.Point(41, 2);
		this.tbFirstWeb.Name = "tbFirstWeb";
		this.tbFirstWeb.ReadOnly = true;
		this.tbFirstWeb.Size = new System.Drawing.Size(889, 20);
		this.tbFirstWeb.TabIndex = 0;
		this.tbFirstWeb.Text = "https://www.drjcertification.org/ter-directory";
		this.panel1.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.panel1.Controls.Add(this.linkLabelFirstWeb);
		this.panel1.Controls.Add(this.tbFirstWeb);
		this.panel1.Location = new System.Drawing.Point(12, 12);
		this.panel1.Name = "panel1";
		this.panel1.Size = new System.Drawing.Size(936, 26);
		this.panel1.TabIndex = 1;
		this.linkLabelFirstWeb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
		this.linkLabelFirstWeb.AutoSize = true;
		this.linkLabelFirstWeb.Location = new System.Drawing.Point(4, 6);
		this.linkLabelFirstWeb.Name = "linkLabelFirstWeb";
		this.linkLabelFirstWeb.Size = new System.Drawing.Size(29, 13);
		this.linkLabelFirstWeb.TabIndex = 1;
		this.linkLabelFirstWeb.TabStop = true;
		this.linkLabelFirstWeb.Text = "TER";
		this.linkLabelFirstWeb.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(linkLabel1_LinkClicked);
		this.panel2.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.panel2.Controls.Add(this.linkLabelSecondWeb);
		this.panel2.Controls.Add(this.tbSecondWeb);
		this.panel2.Location = new System.Drawing.Point(12, 35);
		this.panel2.Name = "panel2";
		this.panel2.Size = new System.Drawing.Size(936, 22);
		this.panel2.TabIndex = 2;
		this.linkLabelSecondWeb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
		this.linkLabelSecondWeb.AutoSize = true;
		this.linkLabelSecondWeb.Location = new System.Drawing.Point(4, 5);
		this.linkLabelSecondWeb.Name = "linkLabelSecondWeb";
		this.linkLabelSecondWeb.Size = new System.Drawing.Size(22, 13);
		this.linkLabelSecondWeb.TabIndex = 1;
		this.linkLabelSecondWeb.TabStop = true;
		this.linkLabelSecondWeb.Text = "ER";
		this.linkLabelSecondWeb.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(linkLabel2_LinkClicked);
		this.tbSecondWeb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.tbSecondWeb.Location = new System.Drawing.Point(41, 2);
		this.tbSecondWeb.Name = "tbSecondWeb";
		this.tbSecondWeb.ReadOnly = true;
		this.tbSecondWeb.Size = new System.Drawing.Size(889, 20);
		this.tbSecondWeb.TabIndex = 0;
		this.tbSecondWeb.Text = "https://www.iapmoes.org/uniform/building-products-evaluation-program/evaluation-report-directory";
		this.tbConsole.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.tbConsole.Location = new System.Drawing.Point(6, 471);
		this.tbConsole.Multiline = true;
		this.tbConsole.Name = "tbConsole";
		this.tbConsole.ReadOnly = true;
		this.tbConsole.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
		this.tbConsole.Size = new System.Drawing.Size(671, 136);
		this.tbConsole.TabIndex = 0;
		this.tbConsole.TextChanged += new System.EventHandler(tbConsole_TextChanged);
		this.tabPage10.Controls.Add(this.gbTools);
		this.tabPage10.Controls.Add(this.gbWebOptions);
		this.tabPage10.Location = new System.Drawing.Point(4, 22);
		this.tabPage10.Name = "tabPage10";
		this.tabPage10.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage10.Size = new System.Drawing.Size(934, 360);
		this.tabPage10.TabIndex = 9;
		this.tabPage10.Text = "Setting";
		this.tabPage10.UseVisualStyleBackColor = true;
		this.gbTools.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.gbTools.Controls.Add(this.panel5);
		this.gbTools.Location = new System.Drawing.Point(432, 7);
		this.gbTools.Name = "gbTools";
		this.gbTools.Size = new System.Drawing.Size(479, 316);
		this.gbTools.TabIndex = 4;
		this.gbTools.TabStop = false;
		this.gbTools.Text = "Tool";
		this.panel5.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.panel5.Controls.Add(this.btnAddTab);
		this.panel5.Controls.Add(this.labelField);
		this.panel5.Controls.Add(this.lbField);
		this.panel5.Controls.Add(this.btnCheckLink);
		this.panel5.Controls.Add(this.tbExcelRange);
		this.panel5.Controls.Add(this.labelExcelRange);
		this.panel5.Controls.Add(this.tbExcelFile);
		this.panel5.Controls.Add(this.labelExcelFile);
		this.panel5.Controls.Add(this.btnUpdateTab);
		this.panel5.Controls.Add(this.cbSelectTab);
		this.panel5.Controls.Add(this.labelSelectTab);
		this.panel5.Location = new System.Drawing.Point(6, 19);
		this.panel5.Name = "panel5";
		this.panel5.Size = new System.Drawing.Size(473, 197);
		this.panel5.TabIndex = 3;
		this.btnAddTab.Location = new System.Drawing.Point(85, 155);
		this.btnAddTab.Name = "btnAddTab";
		this.btnAddTab.Size = new System.Drawing.Size(75, 23);
		this.btnAddTab.TabIndex = 10;
		this.btnAddTab.Text = "Add Tab";
		this.btnAddTab.UseVisualStyleBackColor = true;
		this.btnAddTab.Click += new System.EventHandler(btnAddTab_Click);
		this.labelField.AutoSize = true;
		this.labelField.Location = new System.Drawing.Point(4, 79);
		this.labelField.Name = "labelField";
		this.labelField.Size = new System.Drawing.Size(32, 13);
		this.labelField.TabIndex = 9;
		this.labelField.Text = "Field:";
		this.lbField.ContextMenuStrip = this.listFieldMenu;
		this.lbField.FormattingEnabled = true;
		this.lbField.Items.AddRange(new object[3] { "IAPMO UES ER", "ICC-ES ESR", "LADBS RR" });
		this.lbField.Location = new System.Drawing.Point(86, 79);
		this.lbField.Name = "lbField";
		this.lbField.ScrollAlwaysVisible = true;
		this.lbField.Size = new System.Drawing.Size(171, 69);
		this.lbField.TabIndex = 8;
		this.listFieldMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
		this.listFieldMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[2] { this.toolStripMenuItem1, this.toolStripMenuItem2 });
		this.listFieldMenu.Name = "listFieldMenu";
		this.listFieldMenu.Size = new System.Drawing.Size(118, 48);
		this.listFieldMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(listFieldMenu_ItemClicked);
		this.toolStripMenuItem1.Name = "toolStripMenuItem1";
		this.toolStripMenuItem1.Size = new System.Drawing.Size(117, 22);
		this.toolStripMenuItem1.Text = "Add";
		this.toolStripMenuItem2.Name = "toolStripMenuItem2";
		this.toolStripMenuItem2.Size = new System.Drawing.Size(117, 22);
		this.toolStripMenuItem2.Text = "Remove";
		this.btnCheckLink.Location = new System.Drawing.Point(348, 51);
		this.btnCheckLink.Name = "btnCheckLink";
		this.btnCheckLink.Size = new System.Drawing.Size(73, 23);
		this.btnCheckLink.TabIndex = 7;
		this.btnCheckLink.Text = "Check Link";
		this.btnCheckLink.UseVisualStyleBackColor = true;
		this.tbExcelRange.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.tbExcelRange.Location = new System.Drawing.Point(85, 28);
		this.tbExcelRange.Name = "tbExcelRange";
		this.tbExcelRange.Size = new System.Drawing.Size(382, 20);
		this.tbExcelRange.TabIndex = 6;
		this.tbExcelRange.Text = "B1:H200";
		this.labelExcelRange.AutoSize = true;
		this.labelExcelRange.Location = new System.Drawing.Point(4, 30);
		this.labelExcelRange.Name = "labelExcelRange";
		this.labelExcelRange.Size = new System.Drawing.Size(39, 13);
		this.labelExcelRange.TabIndex = 5;
		this.labelExcelRange.Text = "Range";
		this.tbExcelFile.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.tbExcelFile.Location = new System.Drawing.Point(85, 5);
		this.tbExcelFile.Name = "tbExcelFile";
		this.tbExcelFile.Size = new System.Drawing.Size(395, 20);
		this.tbExcelFile.TabIndex = 4;
		this.tbExcelFile.Text = "20240603 Simpson and Competitor Code Report Summary Final (LOCKED).xlsx";
		this.labelExcelFile.AutoSize = true;
		this.labelExcelFile.Location = new System.Drawing.Point(4, 7);
		this.labelExcelFile.Name = "labelExcelFile";
		this.labelExcelFile.Size = new System.Drawing.Size(52, 13);
		this.labelExcelFile.TabIndex = 3;
		this.labelExcelFile.Text = "Excel File";
		this.btnUpdateTab.Location = new System.Drawing.Point(266, 51);
		this.btnUpdateTab.Name = "btnUpdateTab";
		this.btnUpdateTab.Size = new System.Drawing.Size(75, 23);
		this.btnUpdateTab.TabIndex = 2;
		this.btnUpdateTab.Text = "Update";
		this.btnUpdateTab.UseVisualStyleBackColor = true;
		this.btnUpdateTab.Click += new System.EventHandler(btnUpdateTab_Click);
		this.cbSelectTab.FormattingEnabled = true;
		this.cbSelectTab.Items.AddRange(new object[11]
		{
			"SST ESRs & ERs", "MiTek ESRs & ERs", "Hilti ESRs", "Powers ESRs", "ITW ESRs & ERs", "KC Metals ESRs", "SIKA", "LINFORD", "EJOT", "ACS",
			"Other ESRs & ERs"
		});
		this.cbSelectTab.Location = new System.Drawing.Point(86, 52);
		this.cbSelectTab.Name = "cbSelectTab";
		this.cbSelectTab.Size = new System.Drawing.Size(171, 21);
		this.cbSelectTab.TabIndex = 1;
		this.cbSelectTab.Text = "SST ESRs & ERs";
		this.labelSelectTab.AutoSize = true;
		this.labelSelectTab.Location = new System.Drawing.Point(4, 55);
		this.labelSelectTab.Name = "labelSelectTab";
		this.labelSelectTab.Size = new System.Drawing.Size(65, 13);
		this.labelSelectTab.TabIndex = 0;
		this.labelSelectTab.Text = "Select Tab: ";
		this.gbWebOptions.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
		this.gbWebOptions.Controls.Add(this.panel4);
		this.gbWebOptions.Location = new System.Drawing.Point(6, 6);
		this.gbWebOptions.Name = "gbWebOptions";
		this.gbWebOptions.Size = new System.Drawing.Size(420, 317);
		this.gbWebOptions.TabIndex = 3;
		this.gbWebOptions.TabStop = false;
		this.gbWebOptions.Text = "Web Option";
		this.panel4.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.panel4.Controls.Add(this.labelCheckBy);
		this.panel4.Controls.Add(this.rBtnUseLink);
		this.panel4.Controls.Add(this.rBtnUseWeb);
		this.panel4.Location = new System.Drawing.Point(7, 22);
		this.panel4.Name = "panel4";
		this.panel4.Size = new System.Drawing.Size(407, 131);
		this.panel4.TabIndex = 0;
		this.labelCheckBy.AutoSize = true;
		this.labelCheckBy.Location = new System.Drawing.Point(4, 4);
		this.labelCheckBy.Name = "labelCheckBy";
		this.labelCheckBy.Size = new System.Drawing.Size(79, 13);
		this.labelCheckBy.TabIndex = 2;
		this.labelCheckBy.Text = "Check Link By:";
		this.rBtnUseLink.AutoSize = true;
		this.rBtnUseLink.Location = new System.Drawing.Point(7, 49);
		this.rBtnUseLink.Name = "rBtnUseLink";
		this.rBtnUseLink.Size = new System.Drawing.Size(104, 17);
		this.rBtnUseLink.TabIndex = 1;
		this.rBtnUseLink.Text = "Use Current Link";
		this.rBtnUseLink.UseVisualStyleBackColor = true;
		this.rBtnUseWeb.AutoSize = true;
		this.rBtnUseWeb.Checked = true;
		this.rBtnUseWeb.Location = new System.Drawing.Point(7, 23);
		this.rBtnUseWeb.Name = "rBtnUseWeb";
		this.rBtnUseWeb.Size = new System.Drawing.Size(125, 17);
		this.rBtnUseWeb.TabIndex = 0;
		this.rBtnUseWeb.TabStop = true;
		this.rBtnUseWeb.Text = "Download From Web";
		this.rBtnUseWeb.UseVisualStyleBackColor = true;
		this.tabPage11.Location = new System.Drawing.Point(4, 22);
		this.tabPage11.Name = "tabPage11";
		this.tabPage11.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage11.Size = new System.Drawing.Size(934, 360);
		this.tabPage11.TabIndex = 12;
		this.tabPage11.Text = "ACS";
		this.tabPage11.UseVisualStyleBackColor = true;
		this.tabPage12.Location = new System.Drawing.Point(4, 22);
		this.tabPage12.Name = "tabPage12";
		this.tabPage12.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage12.Size = new System.Drawing.Size(934, 360);
		this.tabPage12.TabIndex = 11;
		this.tabPage12.Text = "Other ESRs & ERs";
		this.tabPage12.UseVisualStyleBackColor = true;
		this.tabPage9.Location = new System.Drawing.Point(4, 22);
		this.tabPage9.Name = "tabPage9";
		this.tabPage9.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage9.Size = new System.Drawing.Size(934, 360);
		this.tabPage9.TabIndex = 8;
		this.tabPage9.Text = "EJOT";
		this.tabPage9.UseVisualStyleBackColor = true;
		this.tabPage8.Location = new System.Drawing.Point(4, 22);
		this.tabPage8.Name = "tabPage8";
		this.tabPage8.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage8.Size = new System.Drawing.Size(934, 360);
		this.tabPage8.TabIndex = 7;
		this.tabPage8.Text = "LINFORD";
		this.tabPage8.UseVisualStyleBackColor = true;
		this.tabPage7.Location = new System.Drawing.Point(4, 22);
		this.tabPage7.Name = "tabPage7";
		this.tabPage7.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage7.Size = new System.Drawing.Size(934, 360);
		this.tabPage7.TabIndex = 6;
		this.tabPage7.Text = "SIKA";
		this.tabPage7.UseVisualStyleBackColor = true;
		this.tabPage6.Location = new System.Drawing.Point(4, 22);
		this.tabPage6.Name = "tabPage6";
		this.tabPage6.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage6.Size = new System.Drawing.Size(934, 360);
		this.tabPage6.TabIndex = 5;
		this.tabPage6.Text = "KC Metals ESRs";
		this.tabPage6.UseVisualStyleBackColor = true;
		this.tabPage5.Location = new System.Drawing.Point(4, 22);
		this.tabPage5.Name = "tabPage5";
		this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage5.Size = new System.Drawing.Size(934, 360);
		this.tabPage5.TabIndex = 4;
		this.tabPage5.Text = "ITW ESRs & ERs";
		this.tabPage5.UseVisualStyleBackColor = true;
		this.tabPage4.Location = new System.Drawing.Point(4, 22);
		this.tabPage4.Name = "tabPage4";
		this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage4.Size = new System.Drawing.Size(934, 360);
		this.tabPage4.TabIndex = 3;
		this.tabPage4.Text = "Powers ESRs";
		this.tabPage4.UseVisualStyleBackColor = true;
		this.tabPage3.Location = new System.Drawing.Point(4, 22);
		this.tabPage3.Name = "tabPage3";
		this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage3.Size = new System.Drawing.Size(934, 360);
		this.tabPage3.TabIndex = 2;
		this.tabPage3.Text = "Hilti ESRs";
		this.tabPage3.UseVisualStyleBackColor = true;
		this.tabPage2.Location = new System.Drawing.Point(4, 22);
		this.tabPage2.Name = "tabPage2";
		this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage2.Size = new System.Drawing.Size(934, 360);
		this.tabPage2.TabIndex = 1;
		this.tabPage2.Text = "MiTek ESRs & ERs";
		this.tabPage2.UseVisualStyleBackColor = true;
		this.tabPage1.Location = new System.Drawing.Point(4, 22);
		this.tabPage1.Name = "tabPage1";
		this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
		this.tabPage1.Size = new System.Drawing.Size(934, 360);
		this.tabPage1.TabIndex = 0;
		this.tabPage1.Text = "SST ESRs & ERs";
		this.tabPage1.UseVisualStyleBackColor = true;
		this.mainTableView.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.mainTableView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
		dataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		dataGridViewCellStyle12.BackColor = System.Drawing.SystemColors.Control;
		dataGridViewCellStyle12.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.6f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		dataGridViewCellStyle12.ForeColor = System.Drawing.SystemColors.WindowText;
		dataGridViewCellStyle12.SelectionBackColor = System.Drawing.SystemColors.Highlight;
		dataGridViewCellStyle12.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
		dataGridViewCellStyle12.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
		this.mainTableView.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle12;
		this.mainTableView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.mainTableView.Columns.AddRange(this.Column1, this.Column2, this.Column3, this.Column4, this.Column5, this.Column6, this.Column7, this.Type, this.Status, this.Column8);
		dataGridViewCellStyle21.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
		dataGridViewCellStyle21.BackColor = System.Drawing.SystemColors.Window;
		dataGridViewCellStyle21.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.6f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		dataGridViewCellStyle21.ForeColor = System.Drawing.SystemColors.ControlText;
		dataGridViewCellStyle21.SelectionBackColor = System.Drawing.SystemColors.Highlight;
		dataGridViewCellStyle21.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
		dataGridViewCellStyle21.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
		this.mainTableView.DefaultCellStyle = dataGridViewCellStyle21;
		this.mainTableView.Location = new System.Drawing.Point(6, 101);
		this.mainTableView.Name = "mainTableView";
		dataGridViewCellStyle22.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
		dataGridViewCellStyle22.BackColor = System.Drawing.SystemColors.Control;
		dataGridViewCellStyle22.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.6f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
		dataGridViewCellStyle22.ForeColor = System.Drawing.SystemColors.WindowText;
		dataGridViewCellStyle22.SelectionBackColor = System.Drawing.SystemColors.Highlight;
		dataGridViewCellStyle22.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
		dataGridViewCellStyle22.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
		this.mainTableView.RowHeadersDefaultCellStyle = dataGridViewCellStyle22;
		this.mainTableView.RowHeadersVisible = false;
		this.mainTableView.RowHeadersWidth = 51;
		this.mainTableView.Size = new System.Drawing.Size(942, 364);
		this.mainTableView.TabIndex = 4;
		this.mainTableView.CellContentDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(mainTableView_CellContentDoubleClick);
		this.mainTableView.MouseClick += new System.Windows.Forms.MouseEventHandler(mainTableView_MouseClick);
		dataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		this.Column1.DefaultCellStyle = dataGridViewCellStyle13;
		this.Column1.HeaderText = "Code Report No";
		this.Column1.MinimumWidth = 6;
		this.Column1.Name = "Column1";
		this.Column1.Resizable = System.Windows.Forms.DataGridViewTriState.True;
		this.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
		dataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
		this.Column2.DefaultCellStyle = dataGridViewCellStyle14;
		this.Column2.HeaderText = "Product Category";
		this.Column2.MinimumWidth = 6;
		this.Column2.Name = "Column2";
		dataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
		this.Column3.DefaultCellStyle = dataGridViewCellStyle15;
		this.Column3.HeaderText = "Description";
		this.Column3.MinimumWidth = 6;
		this.Column3.Name = "Column3";
		dataGridViewCellStyle16.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		this.Column4.DefaultCellStyle = dataGridViewCellStyle16;
		this.Column4.HeaderText = "# Products Listed";
		this.Column4.MinimumWidth = 6;
		this.Column4.Name = "Column4";
		dataGridViewCellStyle17.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		this.Column5.DefaultCellStyle = dataGridViewCellStyle17;
		this.Column5.HeaderText = "Latest Code";
		this.Column5.MinimumWidth = 6;
		this.Column5.Name = "Column5";
		dataGridViewCellStyle18.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		this.Column6.DefaultCellStyle = dataGridViewCellStyle18;
		this.Column6.HeaderText = "Issue/Rev Date";
		this.Column6.MinimumWidth = 6;
		this.Column6.Name = "Column6";
		dataGridViewCellStyle19.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		this.Column7.DefaultCellStyle = dataGridViewCellStyle19;
		this.Column7.HeaderText = "Expiration Date";
		this.Column7.MinimumWidth = 6;
		this.Column7.Name = "Column7";
		dataGridViewCellStyle20.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
		this.Type.DefaultCellStyle = dataGridViewCellStyle20;
		this.Type.HeaderText = "Type";
		this.Type.MinimumWidth = 6;
		this.Type.Name = "Type";
		this.Status.HeaderText = "Status";
		this.Status.MinimumWidth = 6;
		this.Status.Name = "Status";
		this.Status.ReadOnly = true;
		this.Column8.HeaderText = "Process";
		this.Column8.Name = "Column8";
		this.tableMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
		this.tableMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[2] { this.checkLink, this.checkValidLink });
		this.tableMenu.Name = "tableMenu";
		this.tableMenu.Size = new System.Drawing.Size(181, 56);
		this.tableMenu.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(tableMenu_ItemClicked);
		this.checkLink.Image = Code_Report.Properties.Resources.Hyperlink;
		this.checkLink.Name = "checkLink";
		this.checkLink.Size = new System.Drawing.Size(180, 26);
		this.checkLink.Text = "Search Code";
		this.checkValidLink.Image = Code_Report.Properties.Resources.Validation;
		this.checkValidLink.Name = "checkValidLink";
		this.checkValidLink.Size = new System.Drawing.Size(180, 26);
		this.checkValidLink.Text = "Check Missing Link";
		this.tabCompany.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.tabCompany.Controls.Add(this.tabPage1);
		this.tabCompany.Controls.Add(this.tabPage2);
		this.tabCompany.Controls.Add(this.tabPage3);
		this.tabCompany.Controls.Add(this.tabPage4);
		this.tabCompany.Controls.Add(this.tabPage5);
		this.tabCompany.Controls.Add(this.tabPage6);
		this.tabCompany.Controls.Add(this.tabPage7);
		this.tabCompany.Controls.Add(this.tabPage8);
		this.tabCompany.Controls.Add(this.tabPage9);
		this.tabCompany.Controls.Add(this.tabPage11);
		this.tabCompany.Controls.Add(this.tabPage12);
		this.tabCompany.Controls.Add(this.tabPage10);
		this.tabCompany.Location = new System.Drawing.Point(6, 79);
		this.tabCompany.Name = "tabCompany";
		this.tabCompany.SelectedIndex = 0;
		this.tabCompany.Size = new System.Drawing.Size(942, 386);
		this.tabCompany.TabIndex = 5;
		this.tabCompany.SelectedIndexChanged += new System.EventHandler(tabCompany_SelectedIndexChanged);
		this.imageList1.ImageStream = (System.Windows.Forms.ImageListStreamer)resources.GetObject("imageList1.ImageStream");
		this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
		this.imageList1.Images.SetKeyName(0, "Report.png");
		this.imageList1.Images.SetKeyName(1, "Hyperlink.png");
		this.imageList1.Images.SetKeyName(2, "Validation.png");
		this.imageList1.Images.SetKeyName(3, "Button.jpg");
		this.tbThirdWeb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.tbThirdWeb.Location = new System.Drawing.Point(41, 2);
		this.tbThirdWeb.Name = "tbThirdWeb";
		this.tbThirdWeb.ReadOnly = true;
		this.tbThirdWeb.Size = new System.Drawing.Size(891, 20);
		this.tbThirdWeb.TabIndex = 0;
		this.tbThirdWeb.Text = "https://icc-es.org/evaluation-report-program/reports-directory/";
		this.linkLabelThirdWeb.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left;
		this.linkLabelThirdWeb.AutoSize = true;
		this.linkLabelThirdWeb.Location = new System.Drawing.Point(4, 6);
		this.linkLabelThirdWeb.Name = "linkLabelThirdWeb";
		this.linkLabelThirdWeb.Size = new System.Drawing.Size(29, 13);
		this.linkLabelThirdWeb.TabIndex = 1;
		this.linkLabelThirdWeb.TabStop = true;
		this.linkLabelThirdWeb.Text = "ESR";
		this.linkLabelThirdWeb.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(linkLabel3_LinkClicked);
		this.panel3.Anchor = System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
		this.panel3.Controls.Add(this.linkLabelThirdWeb);
		this.panel3.Controls.Add(this.tbThirdWeb);
		this.panel3.Location = new System.Drawing.Point(12, 57);
		this.panel3.Name = "panel3";
		this.panel3.Size = new System.Drawing.Size(936, 24);
		this.panel3.TabIndex = 3;
		this.btnCheckCode.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnCheckCode.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
		this.btnCheckCode.ImageIndex = 3;
		this.btnCheckCode.Location = new System.Drawing.Point(839, 471);
		this.btnCheckCode.Name = "btnCheckCode";
		this.btnCheckCode.Size = new System.Drawing.Size(109, 35);
		this.btnCheckCode.TabIndex = 6;
		this.btnCheckCode.Text = "Check Code";
		this.btnCheckCode.UseVisualStyleBackColor = true;
		this.btnCheckCode.Click += new System.EventHandler(btnCheckCodes_Click);
		this.btnLinkUpdated.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnLinkUpdated.BackColor = System.Drawing.Color.FromArgb(0, 192, 0);
		this.btnLinkUpdated.Location = new System.Drawing.Point(683, 471);
		this.btnLinkUpdated.Name = "btnLinkUpdated";
		this.btnLinkUpdated.Size = new System.Drawing.Size(24, 24);
		this.btnLinkUpdated.TabIndex = 7;
		this.btnLinkUpdated.UseVisualStyleBackColor = false;
		this.labelLinkUpdated.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.labelLinkUpdated.AutoSize = true;
		this.labelLinkUpdated.Location = new System.Drawing.Point(713, 476);
		this.labelLinkUpdated.Name = "labelLinkUpdated";
		this.labelLinkUpdated.Size = new System.Drawing.Size(84, 13);
		this.labelLinkUpdated.TabIndex = 8;
		this.labelLinkUpdated.Text = "Link is Updated.";
		this.labelLinkNotFound.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.labelLinkNotFound.AutoSize = true;
		this.labelLinkNotFound.Location = new System.Drawing.Point(713, 499);
		this.labelLinkNotFound.Name = "labelLinkNotFound";
		this.labelLinkNotFound.Size = new System.Drawing.Size(92, 13);
		this.labelLinkNotFound.TabIndex = 10;
		this.labelLinkNotFound.Text = "No Report Found!";
		this.btnLinkNotFound.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnLinkNotFound.BackColor = System.Drawing.Color.FromArgb(255, 128, 0);
		this.btnLinkNotFound.Location = new System.Drawing.Point(683, 494);
		this.btnLinkNotFound.Name = "btnLinkNotFound";
		this.btnLinkNotFound.Size = new System.Drawing.Size(24, 24);
		this.btnLinkNotFound.TabIndex = 9;
		this.btnLinkNotFound.UseVisualStyleBackColor = false;
		this.labelLinkNotValid.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.labelLinkNotValid.AutoSize = true;
		this.labelLinkNotValid.Location = new System.Drawing.Point(713, 522);
		this.labelLinkNotValid.Name = "labelLinkNotValid";
		this.labelLinkNotValid.Size = new System.Drawing.Size(84, 13);
		this.labelLinkNotValid.TabIndex = 12;
		this.labelLinkNotValid.Text = "Link is not Valid!";
		this.btnLinkNotValid.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnLinkNotValid.BackColor = System.Drawing.Color.Red;
		this.btnLinkNotValid.Location = new System.Drawing.Point(683, 517);
		this.btnLinkNotValid.Name = "btnLinkNotValid";
		this.btnLinkNotValid.Size = new System.Drawing.Size(24, 24);
		this.btnLinkNotValid.TabIndex = 11;
		this.btnLinkNotValid.UseVisualStyleBackColor = false;
		this.labelLinkCancelled.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.labelLinkCancelled.AutoSize = true;
		this.labelLinkCancelled.Location = new System.Drawing.Point(713, 544);
		this.labelLinkCancelled.Name = "labelLinkCancelled";
		this.labelLinkCancelled.Size = new System.Drawing.Size(102, 13);
		this.labelLinkCancelled.TabIndex = 14;
		this.labelLinkCancelled.Text = "Report is Cancelled.";
		this.btnLinkCancelled.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnLinkCancelled.BackColor = System.Drawing.Color.Gray;
		this.btnLinkCancelled.Location = new System.Drawing.Point(683, 539);
		this.btnLinkCancelled.Name = "btnLinkCancelled";
		this.btnLinkCancelled.Size = new System.Drawing.Size(24, 24);
		this.btnLinkCancelled.TabIndex = 13;
		this.btnLinkCancelled.UseVisualStyleBackColor = false;
		this.labelLinkCorrect.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.labelLinkCorrect.AutoSize = true;
		this.labelLinkCorrect.Location = new System.Drawing.Point(713, 567);
		this.labelLinkCorrect.Name = "labelLinkCorrect";
		this.labelLinkCorrect.Size = new System.Drawing.Size(114, 13);
		this.labelLinkCorrect.TabIndex = 16;
		this.labelLinkCorrect.Text = "Current Link is Correct.";
		this.btnLinkCorrect.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnLinkCorrect.BackColor = System.Drawing.Color.White;
		this.btnLinkCorrect.Location = new System.Drawing.Point(683, 562);
		this.btnLinkCorrect.Name = "btnLinkCorrect";
		this.btnLinkCorrect.Size = new System.Drawing.Size(24, 24);
		this.btnLinkCorrect.TabIndex = 15;
		this.btnLinkCorrect.UseVisualStyleBackColor = false;
		this.btnUpdateDate.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnUpdateDate.Location = new System.Drawing.Point(839, 512);
		this.btnUpdateDate.Name = "btnUpdateDate";
		this.btnUpdateDate.Size = new System.Drawing.Size(109, 37);
		this.btnUpdateDate.TabIndex = 18;
		this.btnUpdateDate.Text = "Update Date";
		this.btnUpdateDate.UseVisualStyleBackColor = true;
		this.btnUpdateDate.Click += new System.EventHandler(btnUpdateDate_Click);
		this.btnExcelExport.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.btnExcelExport.Location = new System.Drawing.Point(839, 555);
		this.btnExcelExport.Name = "btnExcelExport";
		this.btnExcelExport.Size = new System.Drawing.Size(108, 36);
		this.btnExcelExport.TabIndex = 19;
		this.btnExcelExport.Text = "Export To Excel";
		this.btnExcelExport.UseVisualStyleBackColor = true;
		this.dataGridViewProgressColumn1.HeaderText = "Process";
		this.dataGridViewProgressColumn1.Name = "dataGridViewProgressColumn1";
		this.dataGridViewProgressColumn1.Width = 94;
		this.pBarState.Anchor = System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right;
		this.pBarState.CustomText = null;
		this.pBarState.DisplayStyle = Code_Report.ProgressBarDisplayText.Percentage;
		this.pBarState.Location = new System.Drawing.Point(683, 590);
		this.pBarState.Name = "pBarState";
		this.pBarState.Size = new System.Drawing.Size(144, 16);
		this.pBarState.TabIndex = 17;
		base.AutoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
		base.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		base.ClientSize = new System.Drawing.Size(954, 619);
		base.Controls.Add(this.btnExcelExport);
		base.Controls.Add(this.btnUpdateDate);
		base.Controls.Add(this.pBarState);
		base.Controls.Add(this.labelLinkCorrect);
		base.Controls.Add(this.btnLinkCorrect);
		base.Controls.Add(this.labelLinkCancelled);
		base.Controls.Add(this.btnLinkCancelled);
		base.Controls.Add(this.labelLinkNotValid);
		base.Controls.Add(this.btnLinkNotValid);
		base.Controls.Add(this.labelLinkNotFound);
		base.Controls.Add(this.btnLinkNotFound);
		base.Controls.Add(this.labelLinkUpdated);
		base.Controls.Add(this.btnLinkUpdated);
		base.Controls.Add(this.btnCheckCode);
		base.Controls.Add(this.mainTableView);
		base.Controls.Add(this.tbConsole);
		base.Controls.Add(this.panel3);
		base.Controls.Add(this.panel2);
		base.Controls.Add(this.panel1);
		base.Controls.Add(this.tabCompany);
		base.Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
		base.Name = "mainFrm";
		this.Text = "Code Report";
		base.FormClosing += new System.Windows.Forms.FormClosingEventHandler(mainFrm_FormClosing);
		base.Load += new System.EventHandler(mainFrm_Load);
		base.ResizeBegin += new System.EventHandler(mainFrm_ResizeBegin);
		base.ResizeEnd += new System.EventHandler(mainFrm_ResizeEnd);
		this.panel1.ResumeLayout(false);
		this.panel1.PerformLayout();
		this.panel2.ResumeLayout(false);
		this.panel2.PerformLayout();
		this.tabPage10.ResumeLayout(false);
		this.gbTools.ResumeLayout(false);
		this.panel5.ResumeLayout(false);
		this.panel5.PerformLayout();
		this.listFieldMenu.ResumeLayout(false);
		this.gbWebOptions.ResumeLayout(false);
		this.panel4.ResumeLayout(false);
		this.panel4.PerformLayout();
		((System.ComponentModel.ISupportInitialize)this.mainTableView).EndInit();
		this.tableMenu.ResumeLayout(false);
		this.tabCompany.ResumeLayout(false);
		this.panel3.ResumeLayout(false);
		this.panel3.PerformLayout();
		base.ResumeLayout(false);
		base.PerformLayout();
	}
}
