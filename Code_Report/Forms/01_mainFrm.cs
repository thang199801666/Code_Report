using iTextSharp.text.pdf.qrcode;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Windows.Forms.LinkLabel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Code_Report
{
    public partial class mainFrm : Form
    {
        public string _currentDirectory = Directory.GetCurrentDirectory();
        public string _currentTab;
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
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
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
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Uri.IsWellFormedUriString(textBox1.Text, UriKind.Absolute))
            {
                System.Diagnostics.Process.Start(textBox1.Text);
            }
        }
        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Uri.IsWellFormedUriString(textBox2.Text, UriKind.Absolute))
            {
                System.Diagnostics.Process.Start(textBox2.Text);
            }
        }
        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Uri.IsWellFormedUriString(textBox3.Text, UriKind.Absolute))
            {
                System.Diagnostics.Process.Start(textBox3.Text);
            }
        }
        private void mainFrm_Load(object sender, EventArgs e)
        {
            string targetDir = Directory.GetCurrentDirectory() + "//Data//";
            if (!Directory.Exists(targetDir))
            {
                System.IO.Directory.CreateDirectory(targetDir);
            }
            string binPath = Directory.GetCurrentDirectory() + "//Data//SST ESRs & ERs.bin";
            if (File.Exists(binPath))
            {
                Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
                foreach (var code in codeReports)
                {
                    DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                    row.Cells[0].ToolTipText = code.Value.Link;
                    row.Cells[0].Value = code.Value.Number;
                    row.Cells[1].Value = code.Value.ProductCategory;
                    row.Cells[2].Value = code.Value.Description;
                    row.Cells[3].Value = code.Value.ProductsListed;
                    row.Cells[4].Value = code.Value.LatestCode;
                    row.Cells[5].Value = code.Value.IssueDate;
                    row.Cells[6].Value = code.Value.ExpirationDate;
                    row.Cells[7].Value = code.Value.WebType;
                    dataGridView1.Rows.Add(row);
                }
            }
        }
        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                string link = dataGridView1.CurrentCell.ToolTipText;
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

        private void button1_Click(object sender, EventArgs e)
        {
            ExcelReader xlReader = new ExcelReader();
            string xlPath = Path.Combine(_currentDirectory, "Data", textBox5.Text);
            if (File.Exists(xlPath))
            {
                xlReader.loadWorkBook(xlPath);
                for (int i = 0; i < comboBox1.Items.Count; i++)
                {
                    string currentText = comboBox1.GetItemText(comboBox1.Items[i]);
                    if ( currentText!="Setting")
                    {
                        Dictionary<string, Codes> codeReports = xlReader.getTableByRange(currentText, textBox6.Text);
                        string binPath = _currentDirectory + "//Data//" + currentText + ".bin"; //comboBox1.Text
                        IOData.WriteToBinaryFile(binPath, codeReports);
                        tbConsole.AppendText("Tab: " + currentText + " Has been updated.");
                        tbConsole.AppendText(Environment.NewLine);
                        GC.Collect();
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
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();
            string binPath = Path.Combine(_currentDirectory, "Data", tabCompany.SelectedTab.Text + ".bin");
            if (File.Exists(binPath))
            {
                Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
                foreach (var code in codeReports)
                {
                    DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[0].Clone();
                    row.Cells[0].ToolTipText = code.Value.Link;
                    row.Cells[0].Value = code.Value.Number;
                    row.Cells[1].Value = code.Value.ProductCategory;
                    row.Cells[2].Value = code.Value.Description;
                    row.Cells[3].Value = code.Value.ProductsListed;
                    row.Cells[4].Value = code.Value.LatestCode;
                    row.Cells[5].Value = code.Value.IssueDate;
                    row.Cells[6].Value = code.Value.ExpirationDate;
                    row.Cells[7].Value = code.Value.WebType;
                    dataGridView1.Rows.Add(row);
                }
            }
        }

        private void btnCheckLink_Click(object sender, EventArgs e)
        {
            //string binPath = Path.Combine(_currentDirectory, "Data", tabCompany.SelectedTab.Text + ".bin");
            //if (File.Exists(binPath))
            //{
            //    Dictionary<string, Codes> codeReports = IOData.ReadFromBinaryFile<Dictionary<string, Codes>>(binPath);
            //}

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (Uri.IsWellFormedUriString(dataGridView1.Rows[i].Cells[0].ToolTipText, UriKind.Absolute))
                {
                    dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Red;
                }
            }
        }

        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            int currentCol = dataGridView1.HitTest(e.X, e.Y).ColumnIndex;
            if (e.Button == MouseButtons.Right && currentCol == 0)
            {
                //ContextMenu tableMenu = new ContextMenu();
                //tableMenu.MenuItems.Add(new System.Windows.Forms.MenuItem("Check Code"));
                //int currentMouseOverRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;
                //tableMenu.Show(dataGridView1, new System.Drawing.Point(e.X, e.Y));
                dataGridView1.ClearSelection();
                int currentRow = dataGridView1.HitTest(e.X, e.Y).RowIndex;
                dataGridView1.Rows[currentRow].Selected = true;
                this.tableMenu.Show(dataGridView1, new System.Drawing.Point(e.X, e.Y));
            }
        }

        private void tableMenu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            ToolStripItem item = e.ClickedItem;
            if (item.Text == "Search Code")
            {
                int rowNo = dataGridView1.SelectedRows[0].Index;
                string code = dataGridView1.Rows[rowNo].Cells[0].Value.ToString();
                string webID = dataGridView1.Rows[rowNo].Cells[7].Value.ToString();

                if (webID == "IAPMO UES ER")
                { 
                    WebReader webReader = new WebReader();
                    string outputText = webReader.trackIAPMOESReport(textBox2.Text, code);
                    //if (!outputText.Contains("no report"))
                    //{

                    //}
                    tbConsole.AppendText(outputText);
                    tbConsole.AppendText(Environment.NewLine);
                }
                if (webID == "ICC-ES ESR")
                {
                    WebReader webReader = new WebReader();
                    string outputText = webReader.trackICCESReport(textBox2.Text, code);
                    //if (!outputText.Contains("no report"))
                    //{

                    //}
                    tbConsole.AppendText(outputText);
                    tbConsole.AppendText(Environment.NewLine);
                }
            }
            if (item.Text == "Check Valid Link")
            {
                for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    if (!Uri.IsWellFormedUriString(dataGridView1.Rows[i].Cells[0].ToolTipText, UriKind.Absolute))
                    {
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Red;
                        tbConsole.AppendText("Link of Code: "+ dataGridView1.Rows[i].Cells[0].Value+" Is not Valid!");
                        tbConsole.AppendText(Environment.NewLine);
                    }
                }
            }
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
                        listBox1.Items.Add(value);
                    }
                    else 
                    {
                        MessageBox.Show("Invalid Value!!!");
                    }
                }
            }
            if (item.Text == "Remove")
            {
                System.Windows.Forms.ListBox.SelectedObjectCollection selectedItems = new System.Windows.Forms.ListBox.SelectedObjectCollection(listBox1);
                selectedItems = listBox1.SelectedItems;

                if (listBox1.SelectedIndex != -1)
                {
                    for (int i = selectedItems.Count - 1; i >= 0; i--)
                        listBox1.Items.Remove(selectedItems[i]);
                }
            }
        }

        private void tbConsole_TextChanged(object sender, EventArgs e)
        {
            tbConsole.ScrollToCaret();
        }
    }
}
