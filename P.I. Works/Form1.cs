

using System;
using System.Data;
using System.Windows.Forms;
using EasyXLS;
using System.Collections;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ExcelDocument workbook = new ExcelDocument();
            DataSet ds = workbook.easy_ReadXLSXActiveSheet_AsDataSet("C:\\ss.xlsx");
            DataTable dataTable = ds.Tables[0];

            for (int column = 0; column < dataTable.Columns.Count; column++)
            {
                dataTable.Columns[column].ColumnName = (string)dataTable.Rows[0][column];
            }
            dataTable.Rows.RemoveAt(0);

            ArrayList C_S_IDs = new ArrayList();
            ArrayList C_S_Counts = new ArrayList();
            ArrayList C_IDs = new ArrayList();

            foreach (DataRow row in dataTable.Rows)
            {
                if (row["PLAY_ID"] == System.DBNull.Value) {
                    break;
                }
                if (row["PLAY_TS"].ToString().Substring(0,10).Equals("08.10.2016")) {
                    if (C_IDs.IndexOf(row["CLIENT_ID"]) == -1)
                    {
                        C_IDs.Add(row["CLIENT_ID"]);
                        C_S_Counts.Add(1);
                        C_S_IDs.Add(new ArrayList());
                        ((ArrayList)(C_S_IDs[C_S_IDs.Count - 1])).Add(row["SONG_ID"]);
                    }
                    else
                    {
                        int index = C_IDs.IndexOf(row["CLIENT_ID"]);
                        if (((ArrayList)C_S_IDs[index]).IndexOf(row["SONG_ID"]) == -1) {
                            ((ArrayList)(C_S_IDs[index])).Add(row["SONG_ID"]);
                            C_S_Counts[index] = ((int)(C_S_Counts[index])) + 1;
                        }
                    }
                }
            }

       
                ArrayList C_Counts = new ArrayList();
            ArrayList IDs = new ArrayList();

            for (int j = 0; j < C_S_Counts.Count; j++) {
                if (IDs.IndexOf(C_S_Counts[j]) == -1) {
                    IDs.Add(C_S_Counts[j]);
                    C_Counts.Add(1);
                } else
                {
                    C_Counts[IDs.IndexOf(C_S_Counts[j])] = ((int)(C_Counts[IDs.IndexOf(C_S_Counts[j])])) + 1;
                }
            }
            StreamWriter output = new StreamWriter("output.txt", false);
            
            for (int j = 0; j < IDs.Count; j++) {
                output.Write(IDs[j] + "\t" + C_Counts[j] + "\r\n");
            }

            output.Flush();

            output.Close();

            
            dataGridView1.DataSource = dataTable.DefaultView;
        }
    }
}