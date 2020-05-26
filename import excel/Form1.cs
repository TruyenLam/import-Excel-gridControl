
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace import_excel
{
    public partial class Form1 : Form
    {
       // new ExcelImport() excellmport;
        public Form1()
        {
            InitializeComponent();
        }

        private void textEdit1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtPath.Text = openFileDialog.FileName;
                    string filename= openFileDialog.FileName;

                    var fullFileName = string.Format("{0}\\{1}", Directory.GetCurrentDirectory(), filename);
                    if (!File.Exists(filename))
                    {
                        System.Windows.Forms.MessageBox.Show("File not found");
                        return;
                    }
                    var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", filename);
                    var adapter = new OleDbDataAdapter("select * from [Sheet$]", connectionString);
                    var ds = new DataSet();
                    string tableName = "excelData";
                    adapter.Fill(ds, tableName);
                    DataTable data = ds.Tables[tableName];

                    gridControl1.DataSource = data;
                    
                }
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog { Filter = "Excel 97-2003|*.xls|Excel Workbook|*.xlsx", Multiselect = false, ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    cboSheet.Items.Clear();
                    txtPath.Text = ofd.FileName;
                    AppHelper h = new AppHelper(String.Format("Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0};Extended Properties={1}Excel 8.0;Imex=2;HDR=yes{1}", ofd.FileName, Convert.ToChar(34)));
                    DataTable sdt = h.GetOleDbSchemaTable();
                    if (sdt.Rows.Count < 1)
                        return;
                    try
                    {
                        foreach (DataRow dr in sdt.Rows)
                        {
                            if (!dr["TABLE_NAME"].ToString().EndsWith("_"))
                                cboSheet.Items.Add(dr["TABLE_NAME"].ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void cboSheet_SelectionChangeCommitted(object sender, EventArgs e)
        {
            using (OleDbConnection cn = new OleDbConnection(string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0; Data Source={0};Extended Properties={1}Excel 8.0;Imex=2;HDR=yes{1}", txtPath.Text, Convert.ToChar(34))))
            {
                using (OleDbDataAdapter adapter = new OleDbDataAdapter("select * from [" + cboSheet.SelectedItem.ToString() + "]", cn))
                {
                    try
                    {
                        using (DataTable dt = new DataTable())
                        {
                            adapter.Fill(dt);
                            gridControl1.DataSource = dt;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}
