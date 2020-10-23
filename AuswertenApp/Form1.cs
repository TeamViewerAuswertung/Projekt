using ExcelDataReader;
using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Web;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace AuswertenApp
{
    public partial class Form1 : Form
    {
        DataGridView out1 = new DataGridView();
        public Form1()
        {
            InitializeComponent();
        }

        public void cboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            var test = cboSheet.SelectedItem.ToString();
            System.Data.DataTable dt = tableCollection[test];
            dataGridView1.DataSource = dt;
        }

        public void button2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = tableCollection[cboSheet.SelectedItem.ToString()];
            System.Data.DataTable dt_clone = new System.Data.DataTable();
            dt.Columns.Add("Neu", typeof(String));
            
            DataRow oldRow = null;
            foreach ( DataRow row in dt.Rows)
            {
                if(oldRow != null)
                {
                    if (row["ID"].Equals(oldRow["ID"]) && row["Benutzer"].Equals(oldRow["Benutzer"]) && DateTime.Compare(Convert.ToDateTime(oldRow["Ende"]), Convert.ToDateTime(row["Start"])) > 0)  
                    {
                        // set Start
                        if (DateTime.Compare(Convert.ToDateTime(oldRow["Start"]), Convert.ToDateTime(row["Start"])) >= 0) oldRow["Start"] = row["Start"];

                        // set Ende
                        if (DateTime.Compare(Convert.ToDateTime(oldRow["Ende"]), Convert.ToDateTime(row["Ende"])) <= 0) oldRow["Ende"] = row["Ende"];
                        
                        oldRow["Dauer"] = Convert.ToDateTime(oldRow["Ende"]).Subtract(Convert.ToDateTime(oldRow["Start"])).TotalMinutes;
                        oldRow["Neu"] = "Y";
                        oldRow["Notizen"] = oldRow["Notizen"] + " // " + row["Notizen"];
                        oldRow["Gebühr"] = Convert.ToDouble(oldRow["Gebühr"]) + Convert.ToDouble(row["Gebühr"]);
                    }
                    else
                    {
                        dt_clone.ImportRow(oldRow);
                        oldRow = row;
                    }
                } else
                {
                    oldRow = row;
                }
            }
            foreach (DataGridViewRow Myrow in dataGridView1.Rows)
            {
                if (Myrow.Cells["Neu"].Value == "Y")
                {
                    Myrow.DefaultCellStyle.BackColor = Color.LightBlue;
                }

                dataGridView1.Sort(dataGridView1.Columns["Start"], ListSortDirection.Ascending);
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                dataGridView1.Columns["Neu"].Visible = false;
            }
        }
        
        DataTableCollection tableCollection;
        private void btnBrowse_Click(object sender, EventArgs e)
        {
            string path = "";

            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = " All Files| *.xlsx; *.xls" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtFilename.Text = openFileDialog.FileName;
                    path = txtFilename.Text;
                    using (var stream = File.Open(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                            {
                                ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                            });
                            tableCollection = result.Tables;
                            cboSheet.Items.Clear();
                            foreach (System.Data.DataTable table in tableCollection)
                                cboSheet.Items.Add(table.TableName);
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            saveFileDialog1.InitialDirectory = "C:\\Users\\Kai\\Desktop";
            saveFileDialog1.Title = "Save as Excel File";
            saveFileDialog1.FileName = "Test";
            saveFileDialog1.Filter = "Excel Files (2007)|*.xlsx";

            if (saveFileDialog1.ShowDialog() != DialogResult.Cancel)
            {
                Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelApp.Application.Workbooks.Add(Type.Missing);
                ExcelApp.Columns.ColumnWidth = 20;

                for (int i = 1; i < dataGridView1.Columns.Count; i++)
                {
                    ExcelApp.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count - 1; j++)
                    {
                        ExcelApp.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
                ExcelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog1.FileName.ToString());
                ExcelApp.ActiveWorkbook.Saved = true;
                ExcelApp.Quit();
                MessageBox.Show("Datei erfolgreich gespeichert!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var rechnung = new Rechnungen();
            var path = txtFilename.Text;
            var zweiListen = rechnung.readExcel(path, cboSheet.Items[0].ToString());
            rechnung.addToPDF(zweiListen[1], "C:\\Users\\kwolt\\Desktop\\Rechnungen\\");
            rechnung.updateExcel(zweiListen[0], path); 
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
