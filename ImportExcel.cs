using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Drawing;

namespace HeatLoss
{
    class ImportExcel
    {
        OleDbConnection conn;
        OleDbDataAdapter MyDataAdapter;
        DataTable dt;

        public void ImportTab(DataGridView dgv,String nameSheet)
        {
            String ruta = "";

            try
            {
               
                ruta = Path.Combine(Application.StartupPath,"Template.xlsx");
                 
                conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;data source=" + ruta + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'");
                MyDataAdapter = new OleDbDataAdapter("Select * from [" + nameSheet + "$]", conn);
                dt = new DataTable();
                MyDataAdapter.Fill(dt);
                dgv.DataSource = dt;
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
                MessageBox.Show(ruta);
            }
        }

        public List<string> ListSheetInExcel()
        {
            OleDbConnectionStringBuilder sbConnection = new OleDbConnectionStringBuilder();
            String strExtendedProperties = String.Empty;
            string filePath = Path.Combine(Application.StartupPath, "Template.xlsx");
            sbConnection.DataSource = filePath;
           
                sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
                strExtendedProperties = "Excel 12.0;HDR=Yes;IMEX=1";
            
            sbConnection.Add("Extended Properties", strExtendedProperties);
            List<string> listSheet = new List<string>();
            using (OleDbConnection conn = new OleDbConnection(sbConnection.ToString()))
            {
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))//checks whether row contains '_xlnm#_FilterDatabase' or sheet name(i.e. sheet name always ends with $ sign)
                    {
                        if (drSheet["TABLE_NAME"].ToString().Replace("$", "").Equals("Resume") || drSheet["TABLE_NAME"].ToString().Replace("$", "").Equals("Rooms"))
                        {

                        }
                        else
                        {
                            listSheet.Add(drSheet["TABLE_NAME"].ToString().Replace("$",""));
                        }
                    }
                }
            }
            return listSheet;
        }

        // Metodo SetUpDataGridView 
        public void SetUpDataGridView(DataGridView myDatagrid)
        {
            DataGridViewCellStyle style = myDatagrid.ColumnHeadersDefaultCellStyle;
            style.BackColor = Color.Navy;
            style.ForeColor = Color.White;
            style.Font = new Font(myDatagrid.Font, FontStyle.Bold);

            myDatagrid.EditMode = DataGridViewEditMode.EditOnEnter;
            myDatagrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders;

            myDatagrid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            myDatagrid.CellBorderStyle = DataGridViewCellBorderStyle.Single;
            myDatagrid.GridColor = SystemColors.ActiveBorder;
            myDatagrid.RowHeadersVisible = false;

            myDatagrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            myDatagrid.MultiSelect = false;
            myDatagrid.BackgroundColor = Color.White;
        }
    }
}

