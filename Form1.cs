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
        private void Check_Lot(DataTable dt, int column_number)
        {
            for (int i = 0; i < Lot_List.Count; i++)
            {
                DataRow[] rows = dt.Select("LOT_ID = '" + Lot_List[i] + "'", "LAYER_SEQ ASC, OPERATION_SEQ ASC");
                DataTable Search_Result = rows.CopyToDataTable();
                bool isFault = false;

                //Check Move_In_Time
                for (int j = 1; j < rows.Length; j++)
                {
                    DateTime Early_Prcs = DateTime.Parse(rows[j-1][column_number].ToString());
                    DateTime Late_Prcs = DateTime.Parse(rows[j][column_number].ToString());
                    if (Early_Prcs > Late_Prcs)
                    {
                        isFault = true;
                    }
                }
               if (isFault)
                {
                    StreamWriter SW = new StreamWriter("Lot_Error_"+dt.Columns[column_number].ColumnName.ToString() + "_" + Lot_List[i] + ".csv", true, Encoding.UTF8);
                    foreach (DataColumn dc in dt.Columns)
                    {
                        SW.Write(dc.ColumnName);
                        SW.Write(',');
                    }
                    SW.Write('\n');
                    foreach (DataRow dr in Search_Result.Rows)
                    {
                        for (int dc = 0; dc < Search_Result.Columns.Count; dc++)
                        {
                            SW.Write(dr[dc].ToString());
                            SW.Write(',');
                        }
                        SW.Write('\n');
                    }
                    SW.Close();
                }
            }
        }

        private void Check_Lot_Row(DataTable dt)
        {
            for (int i = 0; i < Lot_List.Count; i++)
            {
                DataRow[] rows = dt.Select("LOT_ID = '" + Lot_List[i] + "'", "LAYER_SEQ ASC, OPERATION_SEQ ASC");
                DataTable Search_Result = rows.CopyToDataTable();
                bool isFault = false;

                //Check Move_In_Time
                for (int j = 0; j < rows.Length; j++)
                {
                    DateTime Check_In_Time = DateTime.Parse(rows[j][8].ToString());
                    DateTime Move_In_Time = DateTime.Parse(rows[j][9].ToString());
                    DateTime Move_Out_Time = DateTime.Parse(rows[j][10].ToString());
                    DateTime Check_Out_Time = DateTime.Parse(rows[j][11].ToString());
                    if (Check_In_Time > Move_In_Time || Move_In_Time > Move_Out_Time || Move_Out_Time > Check_Out_Time)
                    {
                        isFault = true;
                    }
                }
                if (isFault)
                {
                    StreamWriter SW = new StreamWriter("Lot_Error_SelfTime" + Lot_List[i] + ".csv", true, Encoding.UTF8);
                    foreach (DataColumn dc in dt.Columns)
                    {
                        SW.Write(dc.ColumnName);
                        SW.Write(',');
                    }
                    SW.Write('\n');
                    foreach (DataRow dr in Search_Result.Rows)
                    {
                        for (int dc = 0; dc < Search_Result.Columns.Count; dc++)
                        {
                            SW.Write(dr[dc].ToString());
                            SW.Write(',');
                        }
                        SW.Write('\n');
                    }
                    SW.Close();
                }
            }
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
                    Check_Lot(dt, 8);
                    Check_Lot(dt, 9);
                    Check_Lot(dt, 10);
                    Check_Lot(dt, 11);
                    Check_Lot_Row(dt);
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
