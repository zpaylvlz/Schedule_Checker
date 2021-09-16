using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office;

namespace Schedule_Checker
{
    public partial class Form1 : Form
    {
        Button BtnSelect;
        DataGridView DataVisualize;
        List<string> Lot_List = new List<string>();
        List<string> EQP_List = new List<string>();
        public Form1()
        {
            InitializeComponent();
        }

        private void Check_Lot(DataTable dt)
        {
            for (int i = 0; i < Lot_List.Count; i++)
            {
                DataRow[] rows = dt.Select("LOT_ID = '" + Lot_List[i] + "'", "LAYER_SEQ ASC, OPERATION_SEQ ASC");
                DataTable Search_Result = rows.CopyToDataTable();
                DateTime Early_Prcs = DateTime.Parse(rows[0][9].ToString());
                for (int j = 1; j < rows.Length; j++)
                {
                    DateTime Late_Prcs = DateTime.Parse(rows[j][9].ToString());
                    if (Early_Prcs > Late_Prcs)
                    {
                        MessageBox.Show("Lot Time Invalid");
                    }
                    Early_Prcs = Late_Prcs;
                }
            }
            MessageBox.Show("Done");
            
        }

        private void Check_EQP(DataTable dt)
        {
            for (int i = 0; i < EQP_List.Count; i++)
            {
                DataRow[] rows = dt.Select("TO_EQP_ID = '" + EQP_List[i] + "'", "MOVE_IN_TIME ASC");
                DataTable Search_Result = rows.CopyToDataTable();
                
                for (int j = 1; j < rows.Length; j++)
                {
                    DateTime Last_Prcs_Move_Out = DateTime.Parse(rows[j-1][10].ToString());
                    DateTime Current_Prcs_Move_In = DateTime.Parse(rows[j][9].ToString());
                    if (Last_Prcs_Move_Out > Current_Prcs_Move_In)
                    {
                        MessageBox.Show("EQP Time Invalid");
                    }
                }
            }
            MessageBox.Show("Done");
        }

        private void Select_File(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            OpenFileDialog OF = new OpenFileDialog();
            OF.Filter = "csv file|*.csv;";
            if (OF.ShowDialog() != DialogResult.Cancel)
            {
                
                using (StreamReader sr = new StreamReader(OF.FileName.ToString()))
                {
                    string[] headers = sr.ReadLine().Split(',');
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }
                    dt.Columns[6].DataType = typeof(Int32);
                    dt.Columns[7].DataType = typeof(Int32);
                    dt.Columns[8].DataType = typeof(DateTime);
                    dt.Columns[9].DataType = typeof(DateTime);
                    dt.Columns[10].DataType = typeof(DateTime);
                    dt.Columns[11].DataType = typeof(DateTime);
                    while (!sr.EndOfStream)
                    {
                        string[] rows = sr.ReadLine().Split(',');
                        DataRow dr = dt.NewRow();
                        //Lot index: 2, EQP index: 4;
                        bool Lot_exist = Lot_List.Any(s => s == rows[2]);
                        if (!Lot_exist)
                            Lot_List.Add(rows[2]);
                        bool EQP_exist = EQP_List.Any(s => s == rows[4]);
                        if (!EQP_exist)
                            EQP_List.Add(rows[4]);

                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = rows[i];
                        }
                        dt.Rows.Add(dr);
                    }
                    //DataVisualize.DataSource = dt;
                    Check_Lot(dt);
                    Check_EQP(dt);
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            BtnSelect = new Button();
            BtnSelect.Location = new Point(10, 10);
            BtnSelect.Text = "Select File";
            BtnSelect.AutoSize = true;
            BtnSelect.Click += Select_File;
            this.Controls.Add(BtnSelect);

            DataVisualize = new DataGridView();
            DataVisualize.Size = new Size(1200, 700);
            DataVisualize.Location = new Point(0, 50);
            this.Controls.Add(DataVisualize);
        }
    }
}
