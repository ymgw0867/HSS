using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Globalization;
using HSS.Model;
using Excel = Microsoft.Office.Interop.Excel;

namespace HSS.PrePrint
{
    public partial class frmPrePrint : Form
    {
        int _sStaff;
        
        public frmPrePrint(int sStaff)
        {
            InitializeComponent();
            _sStaff = sStaff;
        }

        private void frmPrePrint_Load(object sender, EventArgs e)
        {

            // フォーム最小値を設定します
            Model.Utility.WindowsMinSize(this, this.Width, this.Height);

            // グリッド定義
            GridviewSet.Setting(dataGridView1, _sStaff);

            // グリッドビューデータ表示
            int rBtnNum = 0;
            if (rBtn1.Checked) rBtnNum = 0;
            else if (rBtn2.Checked) rBtnNum = 1;
            GridviewSet.Show(dataGridView1, _sStaff, rBtnNum);

            // オプション選択欄表示・非表示
            if (_sStaff == global.STAFF_SELECT)
            {
                panel1.Visible = true;
                label7.Visible = true;
                rBtn1.Checked = true;
            }
            else if (_sStaff == global.PART_SELECT)
            {
                panel1.Visible = false;
                label7.Visible = false;
            }

            btnSelAll.Tag = "OFF";

            txtYear.Text = global.sYear.ToString().Substring(2, 2);
            txtMonth.Text = global.sMonth.ToString("00");
        }

        // データグリッドビュークラス
        private class GridviewSet
        {
            public const string col_Check = "col0";
            public const string col_ID = "col1";
            public const string col_Name = "col2";
            public const string col_sID = "col3";
            public const string col_sName1 = "col4";
            public const string col_sName2 = "col5";
            public const string col_Keiyaku = "col6";
            public const string col_WorkTime = "col7";
            public const string col_OrderCode = "col8";
            public const string col_TenpoCode = "col9";

            /// <summary>
            /// データグリッドビューの定義を行います
            /// </summary>
            /// <param name="tempDGV">データグリッドビューオブジェクト</param>
            public static void Setting(DataGridView tempDGV, int sStaff)
            {
                try
                {
                    //フォームサイズ定義

                    // 列スタイルを変更する

                    tempDGV.EnableHeadersVisualStyles = false;

                    // 列ヘッダー表示位置指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                    // 列ヘッダーフォント指定
                    tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                    // データフォント指定
                    tempDGV.DefaultCellStyle.Font = new Font("メイリオ", 9, FontStyle.Regular);

                    // 行の高さ
                    tempDGV.ColumnHeadersHeight = 20;
                    tempDGV.RowTemplate.Height = 20;

                    // 全体の高さ
                    tempDGV.Height = 462;

                    // 奇数行の色
                    //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                    DataGridViewCheckBoxColumn cl = new DataGridViewCheckBoxColumn();

                    //各列幅指定

                    switch (sStaff)
                    {
                        case global.STAFF_SELECT:
                            tempDGV.Columns.Add(col_ID, "個人番号");
                            tempDGV.Columns.Add(col_Name, "氏名");
                            tempDGV.Columns.Add(col_sID, "勤務場所コード");
                            tempDGV.Columns.Add(col_sName1, "勤務場所名称");
                            tempDGV.Columns.Add(col_sName2, "部署");
                            tempDGV.Columns.Add(col_Keiyaku, "契約期間");
                            tempDGV.Columns.Add(col_WorkTime, "勤務時間");
                            tempDGV.Columns.Add(col_OrderCode, "オーダーコード");
                            tempDGV.Columns.Add(col_TenpoCode, "");
                            break;
                        case global.PART_SELECT:
                            tempDGV.Columns.Add(col_ID, "個人番号");
                            tempDGV.Columns.Add(col_Name, "氏名");
                            tempDGV.Columns.Add(col_sID, "勤務場所店番");
                            tempDGV.Columns.Add(col_sName1, "勤務場所部店名");
                            tempDGV.Columns.Add(col_Keiyaku, "契約期間");
                            tempDGV.Columns.Add(col_WorkTime, "勤務時間");
                            break;
                    }

                    tempDGV.Columns[col_ID].Width = 86;
                    tempDGV.Columns[col_Name].Width = 140;
                    tempDGV.Columns[col_sID].Width = 118;
                    tempDGV.Columns[col_sName1].Width = 240;
                    tempDGV.Columns[col_Keiyaku].Width = 200;
                    tempDGV.Columns[col_WorkTime].Width = 100;

                    tempDGV.Columns[col_ID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[col_sID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    
                    if (sStaff == global.STAFF_SELECT)
                    {
                        tempDGV.Columns[col_sName2].Width = 240;
                        tempDGV.Columns[col_OrderCode].Width = 120;
                        tempDGV.Columns[col_TenpoCode].Visible = false;
                        tempDGV.Columns[col_Keiyaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        tempDGV.Columns[col_OrderCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }

                    tempDGV.Columns[col_Keiyaku].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    tempDGV.Columns[col_WorkTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    if (sStaff == global.PART_SELECT) 
                    {
                        tempDGV.Columns[col_sName1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    }

                    // 行ヘッダを表示しない
                    tempDGV.RowHeadersVisible = false;

                    // 選択モード
                    tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    tempDGV.MultiSelect = true;

                    // 編集の可否設定
                    tempDGV.ReadOnly = true;

                    // 追加行表示しない
                    tempDGV.AllowUserToAddRows = false;

                    // データグリッドビューから行削除を禁止する
                    tempDGV.AllowUserToDeleteRows = false;

                    // 手動による列移動の禁止
                    tempDGV.AllowUserToOrderColumns = false;

                    // 列サイズ変更可
                    tempDGV.AllowUserToResizeColumns = true;

                    // 行サイズ変更禁止
                    tempDGV.AllowUserToResizeRows = false;

                    // 行ヘッダーの自動調節
                    //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                    //TAB動作
                    tempDGV.StandardTab = true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ///-----------------------------------------------------------------------------
            /// <summary>
            ///     データグリッドビューにデータを表示します 
            ///     : 西暦対応 2019/02/12 </summary>
            /// <param name="tempGrid">
            ///     データグリッドビューオブジェクト</param>
            /// <param name="usr">
            ///     スタッフ：0, パート：1</param> 
            ///-----------------------------------------------------------------------------
            public static void Show(DataGridView tempGrid, int usr, int rBtn)
            {
                // 2019/02/12 コメント化
                //// 和暦表示用
                //CultureInfo culture = new CultureInfo("ja-JP", true);
                //culture.DateTimeFormat.Calendar = new System.Globalization.JapaneseCalendar();

                // データベース接続
                Model.SysControl.SetDBConnect sDB = new Model.SysControl.SetDBConnect();
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = sDB.cnOpen();
                if (usr == global.STAFF_SELECT)
                {
                    sCom.CommandText = "select * from スタッフマスタ ";
                    if (rBtn == 0)
                    {
                        sCom.CommandText += "where 給与区分 = ? and mid(オーダーコード,1,1) <> ? ";
                    }
                    else if (rBtn == 1)
                    {
                        sCom.CommandText += "where 給与区分 = ? and mid(オーダーコード,1,1) = ? ";
                    }

                    sCom.CommandText += "order by ID";
                    sCom.Parameters.Clear();
                    sCom.Parameters.AddWithValue("@1", "時間給");
                    sCom.Parameters.AddWithValue("@2", "3");

                    OleDbDataReader dr = sCom.ExecuteReader();

                    // マスタ読み込み
                    int iX = 0;
                    tempGrid.RowCount = 0;
                    while (dr.Read())
                    {
                        tempGrid.Rows.Add();
                        tempGrid[col_ID, iX].Value = dr["スタッフコード"].ToString().PadLeft(7, '0');
                        tempGrid[col_Name, iX].Value = dr["スタッフ名"].ToString();
                        tempGrid[col_sID, iX].Value = dr["派遣先CD"].ToString().PadLeft(8, '0');
                        tempGrid[col_sName1, iX].Value = dr["派遣先名"].ToString();
                        tempGrid[col_sName2, iX].Value = dr["派遣先部署"].ToString();
                        tempGrid[col_Keiyaku, iX].Value = dr["契約期間開始"].ToString() + "～" + dr["契約期間終了"].ToString();
                        tempGrid[col_WorkTime, iX].Value = dr["開始時刻1"].ToString().PadLeft(5, '0') + "～" + dr["終了時刻1"].ToString().PadLeft(5, '0');
                        tempGrid[col_OrderCode, iX].Value = dr["オーダーコード"].ToString();
                        tempGrid[col_TenpoCode, iX].Value = dr["派遣先CD"].ToString();

                        iX++;
                    }
                }
                else if (usr == global.PART_SELECT)
                {
                    sCom.CommandText = "select * from パートマスタ order by ID";
                    OleDbDataReader dr = sCom.ExecuteReader();

                    // マスタ読み込み
                    int iX = 0;
                    tempGrid.RowCount = 0;
                    while (dr.Read())
                    {
                        tempGrid.Rows.Add();
                        tempGrid[col_ID, iX].Value = dr["個人番号"].ToString().PadLeft(7, '0');
                        tempGrid[col_Name, iX].Value = dr["姓"].ToString() + " " + dr["名"].ToString();
                        tempGrid[col_sID, iX].Value = dr["勤務場所店番"].ToString().PadLeft(5, '0');
                        tempGrid[col_sName1, iX].Value = dr["勤務場所店名"].ToString();

                        //DateTime d1 = DateTime.Parse(dr["雇用期間始"].ToString()); // 2019/02/12 コメント化
                        //string gd1 = d1.ToString("ggyy/MM/dd", culture);  // 2019/02/12 コメント化
                        //DateTime d2 = DateTime.Parse(dr["雇用期間終"].ToString()); // 2019/02/12 コメント化
                        //string gd2 = d2.ToString("ggyy/MM/dd", culture);  // 2019/02/12 コメント化
                        //tempGrid[col_Keiyaku, iX].Value = gd1 + "～" + gd2; // 2019/02/12 コメント化

                        tempGrid[col_Keiyaku, iX].Value = dr["雇用期間始"].ToString() + "～" + dr["雇用期間終"].ToString();

                        DateTime sTime = DateTime.FromOADate(double.Parse(dr["勤務時間始"].ToString()));
                        DateTime eTime = DateTime.FromOADate(double.Parse(dr["勤務時間終"].ToString()));

                        tempGrid[col_WorkTime, iX].Value = sTime.ToShortTimeString().PadLeft(5,'0') + "～" + eTime.ToShortTimeString().PadLeft(5,'0');
                        iX++;
                    }
                }
                tempGrid.CurrentCell = null;
            }

            ///// <summary>
            ///// データベースのレコードを取得する
            ///// </summary>
            ///// <param name="dgv">グリッドオブジェクト</param>
            ///// <param name="sCode">キーとなる値</param>
            ///// <returns>施工内容マスターインスタンス</returns>
            //public static ms_const GetData(DataGridView dgv, int sCode)
            //{
            //    cnwDataContext db = new cnwDataContext();
            //    ms_const sQuery = db.ms_const.Single(a => a.ID == sCode);
            //    return sQuery;
            //}
        }
        
        private void btnSelAll_Click(object sender, EventArgs e)
        {
            switch (btnSelAll.Tag.ToString())
            {
                case "OFF":
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView1.Rows[i].Selected = true;
                    }
                    btnSelAll.Tag = "ON";
                    btnSelAll.Text = "全て非選択とする(&N)";
                    break;

                case "ON":
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        dataGridView1.Rows[i].Selected = false;
                    }
                    btnSelAll.Tag = "OFF";
                    btnSelAll.Text = "全て選択する(&A)";
                    break;

                default:
                    break;
            }
        }

        private void frmPrePrint_Shown(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
        }

        private void rBtn1_CheckedChanged(object sender, EventArgs e)
        {
            int rBtnNum = 0;
            if (rBtn1.Checked) rBtnNum = 0;
            else if (rBtn2.Checked) rBtnNum = 1;
            GridviewSet.Show(dataGridView1, _sStaff, rBtnNum);
        }

        private void txtSCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\r')
            {
                e.Handled = true;
            }

            if (e.KeyChar == '\r')
            {
                if (txtSCode.Text == string.Empty) return;
                
                dataGridView1.CurrentCell = null;

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1[GridviewSet.col_ID, i].Value.ToString().Contains(txtSCode.Text))
                    {
                        dataGridView1.CurrentCell = dataGridView1[GridviewSet.col_ID, i];
                        break;
                    }
                }
            }
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            // 処理月を検証します
            if (!Utility.NumericCheck(txtMonth.Text))
            {
                MessageBox.Show("処理月が正しくありません","勤務票プレ印刷",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            if (int.Parse(txtMonth.Text) < 1 || int.Parse(txtMonth.Text) > 12)
            {
                MessageBox.Show("正しい処理月を入力してください","勤務票プレ印刷",MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }

            // 印刷対象データ件数を調べます
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("印刷対象データがありません", "勤務票プレ印刷", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string xlsPath = Properties.Settings.Default.instPath + Properties.Settings.Default.XLS + Properties.Settings.Default.ExcelFile;
            if (!System.IO.File.Exists(xlsPath))
            {
                MessageBox.Show("Excelファイルが見つかりません", "勤務票プレ印刷", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("勤務記録用紙の印刷を開始します。よろしいですか？", "勤務票プレ印刷（指定件数：" + dataGridView1.SelectedRows.Count.ToString() + "）"
                , MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                return;

            // 勤務票印刷
            int rBtnNum = 0;
            switch (_sStaff)
            {
                // スタッフ
                case global.STAFF_SELECT:
                    if (rBtn1.Checked == true) rBtnNum = 0;
                    else if (rBtn2.Checked == true) rBtnNum = 1;

                    sReportStaff(Properties.Settings.Default.instPath +
                                 Properties.Settings.Default.XLS +
                                 Properties.Settings.Default.ExcelFile, 
                                 dataGridView1, rBtnNum);
                    break;
                // パートタイマー
                case global.PART_SELECT:
                    sReportPart(Properties.Settings.Default.instPath +
                                 Properties.Settings.Default.XLS +
                                 Properties.Settings.Default.ExcelFile,
                                 dataGridView1);
                    break;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// スタッフ勤務票印刷
        /// </summary>
        /// <param name="xlsPath">勤務票エクセルシートパス</param>
        /// <param name="dg1">データグリッドビューオブジェクト</param>
        /// <param name="rBtn">オプション番号</param>
        private void sReportStaff(string xlsPath, DataGridView dg1, int rBtn)
        {
            string sID = string.Empty;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;
                Excel.Application oXls = new Excel.Application();
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                //Excel.Worksheet oxlsSheet = new Excel.Worksheet();


                // スタッフＩＤによるシートの判断

                string[] a = Properties.Settings.Default.SheetID2.Split(',');
                switch (rBtn)
	            {
		            case 0:
                        a = Properties.Settings.Default.SheetID2.Split(',');
                        break;
                    case 1:
                        a = Properties.Settings.Default.SheetID4.Split(',');
                        break;
	            }

                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[a[0]];
                sID = a[1];
                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                // データベース接続
                SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = sDB.cnOpen();
                OleDbDataReader dr = null;

                try
                {
                    // グリッドを順番に読む
                    for (int i = 0; i < dg1.RowCount; i++)
                    {
                        // 選択されている行を対象とする
                        if (dg1.Rows[i].Selected)
                        {
                            ////印刷2件目以降はシートを追加する
                            //pCnt++;

                            //if (pCnt > 1)
                            //{
                            //    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                            //    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];
                            //}

                            // シートを初期化します
                            oxlsSheet.Cells[3, 2] = string.Empty;
                            oxlsSheet.Cells[3, 5] = string.Empty;
                            oxlsSheet.Cells[3, 6] = string.Empty;
                            oxlsSheet.Cells[3, 7] = string.Empty;
                            oxlsSheet.Cells[3, 8] = string.Empty;
                            oxlsSheet.Cells[3, 9] = string.Empty;
                            oxlsSheet.Cells[3, 10] = string.Empty;
                            oxlsSheet.Cells[3, 11] = string.Empty;
                            oxlsSheet.Cells[3, 14] = string.Empty;
                            oxlsSheet.Cells[3, 26] = string.Empty;
                            oxlsSheet.Cells[3, 27] = string.Empty;
                            oxlsSheet.Cells[3, 28] = string.Empty;
                            oxlsSheet.Cells[3, 29] = string.Empty;
                            oxlsSheet.Cells[3, 31] = string.Empty;
                            oxlsSheet.Cells[3, 32] = string.Empty;
                            oxlsSheet.Cells[4, 4] = string.Empty;
                            oxlsSheet.Cells[4, 26] = string.Empty;
                            oxlsSheet.Cells[5, 4] = string.Empty;
                            oxlsSheet.Cells[5, 5] = string.Empty;
                            oxlsSheet.Cells[5, 6] = string.Empty;
                            oxlsSheet.Cells[5, 7] = string.Empty;
                            oxlsSheet.Cells[5, 8] = string.Empty;
                            oxlsSheet.Cells[5, 9] = string.Empty;
                            oxlsSheet.Cells[5, 10] = string.Empty;
                            oxlsSheet.Cells[5, 11] = string.Empty;
                            oxlsSheet.Cells[5, 14] = string.Empty;
                            oxlsSheet.Cells[5, 28] = string.Empty;
                            oxlsSheet.Cells[5, 29] = string.Empty;
                            oxlsSheet.Cells[5, 30] = string.Empty;
                            oxlsSheet.Cells[5, 31] = string.Empty;
                            oxlsSheet.Cells[5, 32] = string.Empty;
                            oxlsSheet.Cells[5, 33] = string.Empty;
                            oxlsSheet.Cells[5, 34] = string.Empty;

                            for (int ix = 8; ix <= 38; ix++)
                            {
                                oxlsSheet.Cells[ix, 1] = string.Empty;
                                oxlsSheet.Cells[ix, 2] = string.Empty;
                                rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];
                                rng[1] = (Excel.Range)oxlsSheet.Cells[ix, 2];
                                oxlsSheet.get_Range(rng[0], rng[1]).Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                            }
                            
                            int sRow = i;

                            // ID
                            oxlsSheet.Cells[3, 2] = sID;

                            // 個人番号
                            oxlsSheet.Cells[3, 5] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(0, 1);
                            oxlsSheet.Cells[3, 6] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(1, 1);
                            oxlsSheet.Cells[3, 7] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(2, 1);
                            oxlsSheet.Cells[3, 8] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(3, 1);
                            oxlsSheet.Cells[3, 9] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(4, 1);
                            oxlsSheet.Cells[3, 10] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(5, 1);
                            oxlsSheet.Cells[3, 11] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(6, 1);

                            // 氏名
                            oxlsSheet.Cells[3, 14] = dg1[GridviewSet.col_Name, i].Value.ToString(); 

                            // 年
                            oxlsSheet.Cells[3, 26] = "２";
                            oxlsSheet.Cells[3, 27] = "０";
                            oxlsSheet.Cells[3, 28] = this.txtYear.Text.Substring(0, 1);
                            oxlsSheet.Cells[3, 29] = this.txtYear.Text.Substring(1, 1);

                            // 月
                            oxlsSheet.Cells[3, 31] = this.txtMonth.Text.Substring(0, 1);
                            oxlsSheet.Cells[3, 32] = this.txtMonth.Text.Substring(1, 1);

                            // 契約期間
                            oxlsSheet.Cells[4, 4] = dg1[GridviewSet.col_Keiyaku, i].Value.ToString();

                            // 就業時間
                            oxlsSheet.Cells[4, 26] = dg1[GridviewSet.col_WorkTime, i].Value.ToString();

                            // 勤務場所コード
                            oxlsSheet.Cells[5, 4] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(1, 1);
                            oxlsSheet.Cells[5, 5] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(2, 1);
                            oxlsSheet.Cells[5, 6] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(3, 1);
                            oxlsSheet.Cells[5, 7] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(4, 1);
                            oxlsSheet.Cells[5, 8] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(5, 1);
                            oxlsSheet.Cells[5, 9] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(6, 1);
                            oxlsSheet.Cells[5, 10] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(7, 1);
                            oxlsSheet.Cells[5, 11] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(8, 1);

                            // 勤務場所名称
                            if (dg1[GridviewSet.col_TenpoCode, i].Value.ToString().Substring(0,1) == "1")
                                oxlsSheet.Cells[5, 14] = dg1[GridviewSet.col_sName2, i].Value.ToString();
                            else oxlsSheet.Cells[5, 14] = dg1[GridviewSet.col_sName1, i].Value.ToString() + " " + dg1[GridviewSet.col_sName2, i].Value.ToString();

                            // オーダーコード
                            oxlsSheet.Cells[5, 28] = dg1[GridviewSet.col_OrderCode, i].Value.ToString().Substring(0, 1);
                            oxlsSheet.Cells[5, 29] = dg1[GridviewSet.col_OrderCode, i].Value.ToString().Substring(1, 1);
                            oxlsSheet.Cells[5, 30] = dg1[GridviewSet.col_OrderCode, i].Value.ToString().Substring(2, 1);
                            oxlsSheet.Cells[5, 31] = dg1[GridviewSet.col_OrderCode, i].Value.ToString().Substring(3, 1);
                            oxlsSheet.Cells[5, 32] = dg1[GridviewSet.col_OrderCode, i].Value.ToString().Substring(4, 1);
                            oxlsSheet.Cells[5, 33] = dg1[GridviewSet.col_OrderCode, i].Value.ToString().Substring(5, 1);
                            oxlsSheet.Cells[5, 34] = dg1[GridviewSet.col_OrderCode, i].Value.ToString().Substring(6, 1);

                            // 日・曜日
                            DateTime dt;
                            string sDate = "20" + txtYear.Text + "/" + txtMonth.Text + "/";
                            for (int ix = 8; ix <= 38; ix++)
                            {
                                int days = ix - 7;
                                if (DateTime.TryParse(sDate + days.ToString(), out dt))
                                {
                                    string youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                                    oxlsSheet.Cells[ix, 1] = days.ToString();
                                    oxlsSheet.Cells[ix, 2] = youbi; 
                                    rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];

                                    // 土日の場合
                                    if (youbi == "日" || youbi == "土")
                                    {
                                        oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
                                    }
                                    else
                                    {
                                        // 休日テーブルを参照し休日に該当するか調べます
                                        sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                                        sCom.Parameters.Clear();
                                        sCom.Parameters.AddWithValue("@year", int.Parse("20" + txtYear.Text));
                                        sCom.Parameters.AddWithValue("@Month", int.Parse(txtMonth.Text));
                                        sCom.Parameters.AddWithValue("@day", days);
                                        dr = sCom.ExecuteReader();
                                        if (dr.HasRows)
                                        {
                                            oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
                                        }
                                        dr.Close();
                                    }
                                }
                            }
                            // ウィンドウを非表示にする
                            //oXls.Visible = false;
                            // 印刷
                            //oxlsSheet.PrintPreview(false);
                            oxlsSheet.PrintOut(1, Type.Missing, 1, false, oXls.ActivePrinter, Type.Missing, Type.Missing, Type.Missing);
                            //oXlsBook.PrintOut();
                        }
                    }

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 終了メッセージ
                    MessageBox.Show("印刷が終了しました","勤務票印刷",MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                    // COMオブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // データリーダーを閉じる
                    if (dr.IsClosed == false) dr.Close();

                    // データベースを切断
                    if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        /// <summary>
        /// パートタイマー勤務票印刷
        /// </summary>
        /// <param name="xlsPath">勤務票エクセルシートパス</param>
        /// <param name="dg1">データグリッドビューオブジェクト</param>
        private void sReportPart(string xlsPath, DataGridView dg1)
        {
            string sID = string.Empty;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;
                Excel.Application oXls = new Excel.Application();
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                //Excel.Worksheet oxlsSheet = new Excel.Worksheet();

                // スタッフＩＤによるシートの判断
                Excel.Worksheet oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets["ID1"];
                sID = "1";

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                // データベース接続
                SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
                OleDbCommand sCom = new OleDbCommand();
                sCom.Connection = sDB.cnOpen();
                OleDbDataReader dr = null;

                try
                {
                    //グリッドを順番に読む
                    for (int i = 0; i < dg1.RowCount; i++)
                    {
                        //チェックがあるものを対象とする
                        if (dg1.Rows[i].Selected)
                        {
                            ////印刷2件目以降はシートを追加する
                            //pCnt++;

                            //if (pCnt > 1)
                            //{
                            //    oxlsSheet.Copy(Type.Missing, oXlsBook.Sheets[pCnt - 1]);
                            //    oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt];
                            //}

                            // シートを初期化します
                            oxlsSheet.Cells[3, 2] = string.Empty;
                            oxlsSheet.Cells[3, 5] = string.Empty;
                            oxlsSheet.Cells[3, 6] = string.Empty;
                            oxlsSheet.Cells[3, 7] = string.Empty;
                            oxlsSheet.Cells[3, 8] = string.Empty;
                            oxlsSheet.Cells[3, 9] = string.Empty;
                            oxlsSheet.Cells[3, 10] = string.Empty;
                            oxlsSheet.Cells[3, 11] = string.Empty;
                            oxlsSheet.Cells[3, 14] = string.Empty;
                            oxlsSheet.Cells[3, 26] = string.Empty;
                            oxlsSheet.Cells[3, 27] = string.Empty;
                            oxlsSheet.Cells[3, 28] = string.Empty;
                            oxlsSheet.Cells[3, 29] = string.Empty;
                            oxlsSheet.Cells[3, 31] = string.Empty;
                            oxlsSheet.Cells[3, 32] = string.Empty;
                            oxlsSheet.Cells[4, 4] = string.Empty;
                            oxlsSheet.Cells[4, 26] = string.Empty;
                            oxlsSheet.Cells[5, 4] = string.Empty;
                            oxlsSheet.Cells[5, 5] = string.Empty;
                            oxlsSheet.Cells[5, 6] = string.Empty;
                            oxlsSheet.Cells[5, 11] = string.Empty;

                            for (int ix = 8; ix <= 38; ix++)
                            {
                                oxlsSheet.Cells[ix, 1] = string.Empty;
                                oxlsSheet.Cells[ix, 2] = string.Empty;
                                rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];
                                rng[1] = (Excel.Range)oxlsSheet.Cells[ix, 2];
                                oxlsSheet.get_Range(rng[0], rng[1]).Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                            }

                            int sRow = i;

                            // ID
                            oxlsSheet.Cells[3, 2] = sID;

                            // 個人番号
                            oxlsSheet.Cells[3, 5] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(0, 1);
                            oxlsSheet.Cells[3, 6] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(1, 1);
                            oxlsSheet.Cells[3, 7] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(2, 1);
                            oxlsSheet.Cells[3, 8] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(3, 1);
                            oxlsSheet.Cells[3, 9] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(4, 1);
                            oxlsSheet.Cells[3, 10] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(5, 1);
                            oxlsSheet.Cells[3, 11] = dg1[GridviewSet.col_ID, i].Value.ToString().Substring(6, 1);

                            // 氏名
                            oxlsSheet.Cells[3, 14] = dg1[GridviewSet.col_Name, i].Value.ToString();

                            // 年
                            oxlsSheet.Cells[3, 26] = "２";
                            oxlsSheet.Cells[3, 27] = "０";
                            oxlsSheet.Cells[3, 28] = this.txtYear.Text.Substring(0, 1);
                            oxlsSheet.Cells[3, 29] = this.txtYear.Text.Substring(1, 1);

                            // 月
                            oxlsSheet.Cells[3, 31] = this.txtMonth.Text.Substring(0, 1);
                            oxlsSheet.Cells[3, 32] = this.txtMonth.Text.Substring(1, 1);

                            // 契約期間
                            oxlsSheet.Cells[4, 4] = dg1[GridviewSet.col_Keiyaku, i].Value.ToString();

                            // 就業時間
                            oxlsSheet.Cells[4, 26] = dg1[GridviewSet.col_WorkTime, i].Value.ToString();

                            // 勤務場所店番
                            oxlsSheet.Cells[5, 4] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(2, 1);
                            oxlsSheet.Cells[5, 5] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(3, 1);
                            oxlsSheet.Cells[5, 6] = dg1[GridviewSet.col_sID, i].Value.ToString().Substring(4, 1);

                            // 勤務場所名称
                            oxlsSheet.Cells[5, 11] = dg1[GridviewSet.col_sName1, i].Value.ToString();

                            // 日・曜日
                            DateTime dt;
                            string sDate = "20" + txtYear.Text + "/" + txtMonth.Text + "/";
                            for (int ix = 8; ix <= 38; ix++)
                            {
                                int days = ix - 7;
                                if (DateTime.TryParse(sDate + days.ToString(), out dt))
                                {
                                    string youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                                    oxlsSheet.Cells[ix, 1] = days.ToString();
                                    oxlsSheet.Cells[ix, 2] = youbi;
                                    rng[0] = (Excel.Range)oxlsSheet.Cells[ix, 1];

                                    // 土日の場合
                                    if (youbi == "日" || youbi == "土")
                                    {
                                        oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
                                    }
                                    else
                                    {
                                        // 休日テーブルを参照し休日に該当するか調べます
                                        sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                                        sCom.Parameters.Clear();
                                        sCom.Parameters.AddWithValue("@year", int.Parse("20" + txtYear.Text));
                                        sCom.Parameters.AddWithValue("@Month", int.Parse(txtMonth.Text));
                                        sCom.Parameters.AddWithValue("@day", days);
                                        dr = sCom.ExecuteReader();
                                        if (dr.HasRows)
                                        {
                                            oxlsSheet.get_Range(rng[0], rng[0]).Interior.ColorIndex = 15;
                                        }
                                        dr.Close();
                                    }
                                }
                            }
                            // ウィンドウを非表示にする
                            //oXls.Visible = false;
                            //印刷
                            //oxlsSheet.PrintPreview(false);
                            oxlsSheet.PrintOut(1, Type.Missing, 1, false, oXls.ActivePrinter, Type.Missing, Type.Missing, Type.Missing);
                            //oXlsBook.PrintOut();
                        }
                    }

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;
                    // 終了メッセージ
                    MessageBox.Show("印刷が終了しました", "勤務票印刷", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }

                finally
                {
                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    // 保存処理
                    oXls.DisplayAlerts = false;

                    // Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // Excelを終了
                    oXls.Quit();

                    // COMオブジェクトの参照カウントを解放する 
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                    // マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // データリーダーを閉じる
                    if (dr.IsClosed == false) dr.Close();

                    // データベースを切断
                    if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            // マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }

        private void txtMonth_Leave(object sender, EventArgs e)
        {
            txtMonth.Text = txtMonth.Text.PadLeft(2, '0');
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            txtYear.Text = txtYear.Text.PadLeft(2, '0');
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmPrePrint_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
