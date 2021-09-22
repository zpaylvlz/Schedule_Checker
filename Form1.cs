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

namespace Schedule_Checker
{
    public partial class Form1 : Form
    {
        Button BtnSelect;
        DataGridView DataVisualize;
        List<string> Lot_List;
        List<string> EQP_List;
        public Form1()
        {
            InitializeComponent();
        }
        private void Check_Lot(DataTable dt, int column_number)
        {
            List<string> Error_Seq = new List<string>();
            for (int i = 0; i < Lot_List.Count; i++)
            {
                DataRow[] rows = dt.Select("LOT_ID = '" + Lot_List[i] + "'", "LAYER_SEQ ASC, OPERATION_SEQ ASC");
                DataTable Search_Result = rows.CopyToDataTable();
                bool isFault = false;

                //Check Time By Selected column
                
                for (int j = 1; j < rows.Length; j++)
                {
                    DateTime Early_Prcs = DateTime.Parse(rows[j-1][column_number].ToString());
                    DateTime Late_Prcs = DateTime.Parse(rows[j][column_number].ToString());
                    int Date_Compare = DateTime.Compare(Early_Prcs, Late_Prcs);
                    if (Date_Compare > 0)
                    {
                        Error_Seq.Add("Layer_SEQ: " + rows[j - 1][6].ToString() + "& Operation_SEQ: " + rows[j - 1][7].ToString());
                        isFault = true;
                    }
                }
               if (isFault)
                {
                    StreamWriter SW = new StreamWriter("Lot_Error_"+dt.Columns[column_number].ColumnName.ToString() + "_" + Lot_List[i] + ".csv", false, Encoding.UTF8);
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
                    foreach(string s in Error_Seq)
                    {
                        SW.WriteLine(s);
                    }
                    SW.Close();
                }
            }
        }

        private void Check_Lot_Row(DataTable dt)
        {
            for (int i = 0; i < Lot_List.Count; i++)
            {
                List<string> Error_Seq = new List<string>();
                DataRow[] rows = dt.Select("LOT_ID = '" + Lot_List[i] + "'", "LAYER_SEQ ASC, OPERATION_SEQ ASC");
                DataTable Search_Result = rows.CopyToDataTable();
                bool isFault = false;

                //Check Time Sequence of a row
                for (int j = 0; j < rows.Length; j++)
                {
                    DateTime Check_In_Time = DateTime.Parse(rows[j][8].ToString());
                    DateTime Move_In_Time = DateTime.Parse(rows[j][9].ToString());
                    DateTime Move_Out_Time = DateTime.Parse(rows[j][10].ToString());
                    DateTime Check_Out_Time = DateTime.Parse(rows[j][11].ToString());
                    int Date_Compare_Check2Move = DateTime.Compare(Check_In_Time, Move_In_Time);
                    int Date_Compare_MoveInOut = DateTime.Compare(Move_In_Time, Move_Out_Time);
                    int Date_Compare_Move2Check = DateTime.Compare(Move_Out_Time, Check_Out_Time);
                    if (Date_Compare_Check2Move > 0 || Date_Compare_MoveInOut > 0 || Date_Compare_Move2Check > 0)
                    {
                        Error_Seq.Add("Layer_SEQ: " + rows[j][6].ToString() + "& Operation_SEQ: " + rows[j][7].ToString());
                        isFault = true;
                    }
                }
                if (isFault)
                {
                    StreamWriter SW = new StreamWriter("Lot_Error_SelfTime_" + Lot_List[i] + ".csv", false, Encoding.UTF8);
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
                    foreach (string s in Error_Seq)
                    {
                        SW.WriteLine(s);
                    }
                    SW.Close();
                }
            }
        }

        private void Check_EQP(DataTable dt)
        {
            //DataRow[] rows = dt.Select("TO_EQP_ID = '" + EQP_List[0] + "'", "Move_IN_TIME ASC");
            //DataTable Search_Result = rows.CopyToDataTable();
            //DataVisualize.DataSource = Search_Result;
             for (int i = 0; i < EQP_List.Count; i++)
             {
                    DataRow[] rows = dt.Select("TO_EQP_ID = '" + EQP_List[i] + "'", "MOVE_IN_TIME ASC");
                    DataTable Search_Result = rows.CopyToDataTable();
                    bool isFault = false;
                    List<string> Error_Seq = new List<string>();
                    for (int j = 1; j < rows.Length; j++)
                    {
                        DateTime Last_Prcs_Move_In = DateTime.Parse(rows[j - 1][9].ToString());
                        DateTime Last_Prcs_Move_Out = DateTime.Parse(rows[j - 1][10].ToString());
                        DateTime Current_Prcs_Move_In = DateTime.Parse(rows[j][9].ToString());

                        int Date_Compare_Queue = DateTime.Compare(Last_Prcs_Move_Out, Current_Prcs_Move_In);
                        int Date_Compare_In = DateTime.Compare(Last_Prcs_Move_In, Current_Prcs_Move_In);
                        if (Date_Compare_Queue > 0 || Date_Compare_In > 0)
                        {
                            isFault = true;
                            Error_Seq.Add("Index: " + (j-1).ToString() + " & Layer_SEQ: " + rows[j - 1][6].ToString() + " & Operation_SEQ: " + rows[j - 1][7].ToString());
                        }

                        #region write file
                        if (isFault)
                        {
                            StreamWriter SW = new StreamWriter("EQP_Error_Time_Overlap_" + EQP_List[i] + ".csv", false, Encoding.UTF8);
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
                            foreach (string s in Error_Seq)
                            {
                                SW.WriteLine(s);
                            }
                            SW.Close();
                        }
                        #endregion
                    }
             }
        }

        private void Check_Early_Enter(DataTable dt)
        {
            //Check if there's any process check in earlier but move in later then next process
            for (int i = 0; i < EQP_List.Count; i++)
            {
                DataRow[] rows = dt.Select("TO_EQP_ID = '" + EQP_List[i] + "'", "MOVE_IN_TIME ASC");
                DataTable Search_Result = rows.CopyToDataTable();
                bool isFault = false;
                List<string> Error_Seq = new List<string>();
                for (int j = 1; j < rows.Length; j++)
                {
                    DateTime Last_Prcs_Check_In = DateTime.Parse(rows[j - 1][8].ToString());
                    DateTime Current_Prcs_Check_In = DateTime.Parse(rows[j][8].ToString());

                    int Date_Compare_Queue = DateTime.Compare(Last_Prcs_Check_In, Current_Prcs_Check_In);
                    if (Date_Compare_Queue > 0)
                    {
                        isFault = true;
                        Error_Seq.Add("Index: " + (j - 1).ToString() + " & Layer_SEQ: " + rows[j - 1][6].ToString() + " & Operation_SEQ: " + rows[j - 1][7].ToString());
                    }

                    #region write file
                    if (isFault)
                    {
                        StreamWriter SW = new StreamWriter("EQP_Check_In_Earlier_" + EQP_List[i] + ".csv", false, Encoding.UTF8);
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
                        foreach (string s in Error_Seq)
                        {
                            SW.WriteLine(s);
                        }
                        SW.Close();
                    }
                    #endregion
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
                    dt = new DataTable();
                    Lot_List = new List<string>();
                    EQP_List = new List<string>();
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
                    Check_Lot(dt, 8);
                    Check_Lot(dt, 9);
                    Check_Lot(dt, 10);
                    Check_Lot(dt, 11);
                    Check_Lot_Row(dt);
                    Check_EQP(dt);
                    Check_Early_Enter(dt);
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
