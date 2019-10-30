using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using HSS.Model;

namespace HSS.ScanOcr
{
    public partial class frmStaffSelect : Form
    {
        public frmStaffSelect(int usr)
        {
            InitializeComponent();
            _usrSel = usr;
            _msCode = string.Empty;
        }

        private void frmStaffSelect_Load(object sender, EventArgs e)
        {
            Model.Utility.WindowsMaxSize(this, this.Width, this.Height);
            Model.Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビュー定義
            if (_usrSel == global.STAFF_SELECT)
            {
                GridViewStaffSetting(dataGridView1);    // スタッフ定義
                GridViewStaffShow(dataGridView1);       // スタッフマスター表示
            }
            else if (_usrSel == global.PART_SELECT)
            {
                GridViewPartSetting(dataGridView1);     // パートタイマー定義
                GridViewPartShow(dataGridView1);        // パートタイマーマスター表示
            }
        }

        // 指定スタッフ
        private int _usrSel;

        // データグリッドビューカラム名
        private string cCode = "c1";
        private string cName = "c2";
        private string cKinmuCode = "c3";
        private string cKinmuName = "c4";
        private string cKinmuBusho = "c5";
        private string cKKikan = "c6";
        private string cKTime = "c7";
        private string cKOrderCode = "c8";
        private string cTcode = "c9";

        /// <summary>
        /// スタッフマスターグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewStaffSetting(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 380;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cCode, "個人番号");
                tempDGV.Columns.Add(cName, "氏名");
                tempDGV.Columns.Add(cKinmuCode, "勤務場所コード");
                tempDGV.Columns.Add(cKinmuName, "勤務場所名称");
                tempDGV.Columns.Add(cKinmuBusho, "部署");
                tempDGV.Columns.Add(cKKikan, "契約期間");
                tempDGV.Columns.Add(cKTime, "勤務時間");
                tempDGV.Columns.Add(cKOrderCode, "オーダーコード");
                tempDGV.Columns.Add(cTcode, "コード");
                tempDGV.Columns[cTcode].Visible = false;

                tempDGV.Columns[cCode].Width = 80;
                tempDGV.Columns[cName].Width = 120;
                tempDGV.Columns[cKinmuCode].Width = 120;
                tempDGV.Columns[cKinmuName].Width = 200;
                tempDGV.Columns[cKinmuBusho].Width = 160;
                tempDGV.Columns[cKKikan].Width = 200;
                tempDGV.Columns[cKTime].Width = 120;
                tempDGV.Columns[cKOrderCode].Width = 100;

                tempDGV.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKinmuCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKKikan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKOrderCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                
                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                tempDGV.StandardTab = false;

                // ソート禁止
                foreach (DataGridViewColumn c in tempDGV.Columns)
                {
                    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                //tempDGV.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// スタッフマスターグリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewPartSetting(DataGridView tempDGV)
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
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 380;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cCode, "個人番号");
                tempDGV.Columns.Add(cName, "氏名");
                tempDGV.Columns.Add(cKinmuCode, "勤務場所店番");
                tempDGV.Columns.Add(cKinmuName, "勤務場所部店名");
                tempDGV.Columns.Add(cKKikan, "契約期間");
                tempDGV.Columns.Add(cKTime, "勤務時間");

                tempDGV.Columns[cCode].Width = 80;
                tempDGV.Columns[cName].Width = 120;
                tempDGV.Columns[cKinmuCode].Width = 120;
                tempDGV.Columns[cKinmuName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                tempDGV.Columns[cKKikan].Width = 200;
                tempDGV.Columns[cKTime].Width = 120;

                tempDGV.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKinmuCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKKikan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cKTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // 編集不可とする
                tempDGV.ReadOnly = true;

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更不可
                tempDGV.AllowUserToResizeColumns = false;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                tempDGV.StandardTab = false;

                // ソート禁止
                foreach (DataGridViewColumn c in tempDGV.Columns)
                {
                    c.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                //tempDGV.Columns[cDay].SortMode = DataGridViewColumnSortMode.NotSortable;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        /// <summary>
        /// スタッフマスタをグリッドビューへ表示します
        /// </summary>
        /// <param name="tempGrid">データグリッドビューオブジェクト</param>
        private void GridViewStaffShow(DataGridView tempGrid)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            string mySql = "select * from スタッフマスタ order by スタッフコード";
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = mySql;
            dr = sCom.ExecuteReader();

            int iX = 0;
            tempGrid.RowCount = 0;

            while (dr.Read())
            {
                tempGrid.Rows.Add();

                tempGrid[cCode, iX].Value = dr["スタッフコード"].ToString();
                tempGrid[cName, iX].Value = dr["スタッフ名"].ToString();
                tempGrid[cKinmuCode, iX].Value = dr["派遣先CD"].ToString().Substring(1, 8);
                tempGrid[cKinmuName, iX].Value = dr["派遣先名"].ToString();
                tempGrid[cKinmuBusho, iX].Value = dr["派遣先部署"].ToString();
                tempGrid[cKKikan, iX].Value = dr["契約期間開始"].ToString() + "～" + dr["契約期間終了"].ToString();
                string st = dr["開始時刻1"].ToString().PadLeft(5, '0');
                string et = dr["終了時刻1"].ToString().PadLeft(5, '0'); 
                tempGrid[cKTime, iX].Value = st + "～" + et;
                tempGrid[cKOrderCode, iX].Value = dr["オーダーコード"].ToString();
                tempGrid[cTcode, iX].Value = dr["派遣先CD"].ToString();

                iX++;
            }

            dr.Close();
            sCom.Connection.Close();

            tempGrid.CurrentCell = null;  
        }

        /// <summary>
        /// パートタイマーをグリッドビューへ表示します
        /// </summary>
        /// <param name="tempGrid">データグリッドビューオブジェクト</param>
        private void GridViewPartShow(DataGridView tempGrid)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            string mySql = "select * from パートマスタ order by 個人番号";
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = mySql;
            dr = sCom.ExecuteReader();

            int iX = 0;
            tempGrid.RowCount = 0;

            while (dr.Read())
            {
                tempGrid.Rows.Add();

                tempGrid[cCode, iX].Value = dr["個人番号"].ToString();
                tempGrid[cName, iX].Value = dr["姓"].ToString() + " " + dr["名"].ToString();
                tempGrid[cKinmuCode, iX].Value = dr["勤務場所店番"].ToString();
                tempGrid[cKinmuName, iX].Value = dr["勤務場所店名"].ToString();
                tempGrid[cKKikan, iX].Value = dr["雇用期間始"].ToString() + "～" + dr["雇用期間終"].ToString();
                string st = DateTime.FromOADate(double.Parse(dr["勤務時間始"].ToString())).ToShortTimeString().PadLeft(5, '0');
                string et = DateTime.FromOADate(double.Parse(dr["勤務時間終"].ToString())).ToShortTimeString().PadLeft(5, '0');
                tempGrid[cKTime, iX].Value = st + "～" + et;

                iX++;
            }

            dr.Close();
            sCom.Connection.Close();

            tempGrid.CurrentCell = null;
        }

        private void GetGridViewData(DataGridView g)
        {
            if (g.SelectedRows.Count == 0) return;

            int r = g.SelectedRows[0].Index;

            _msCode = g[cCode, r].Value.ToString();

            //_msName = g[cName, r].Value.ToString();
            //_msTenpoCode = g[cKinmuCode, r].Value.ToString();
            //_msTenpoName = g[cKinmuName, r].Value.ToString();
            //if (_usrSel == global.STAFF_SELECT)
            //{
            //    _msBusho = g[cKinmuBusho, r].Value.ToString();
            //    _msTcode = g[cTcode, r].Value.ToString();
            //    _msOrderCode = g[cKOrderCode, r].Value.ToString();
            //}

            //_msStime = g[cKKikan, r].Value.ToString().Substring(0, 5);
            //_msEtime = g[cKKikan, r].Value.ToString().Substring(6, 5);
                   
            //TimeS = Replace(msStime, ":", "");
            //TimeE = Replace(msEtime, ":", "");
            //TimeJ = (((fncCMin(TimeE) - fncCMin(TimeS)) - 60) / 2);
    
            //_msHStime1 = msStime;
            //_msHEtime1 = Format(DateAdd("n", TimeJ, msStime), "Short Time");
            //_msHStime2 = Format(DateAdd("n", TimeJ + 60, msStime), "Short Time");
            //_msHEtime2 = msEtime;
        }

        public string _msCode { get; set; }
        public string _msName { get; set; }
        public string _msTenpoCode { get; set; }
        public string _msTenpoName { get; set; }
        public string _msBusho { get; set; }
        public string _msTcode { get; set; }
        public string _msOrderCode { get; set; }
        public string _msStime { get; set; }
        public string _msEtime { get; set; }
        public string _msHStime1 { get; set; }
        public string _msHEtime1 { get; set; }
        public string _msHStime2 { get; set; }
        public string _msHEtime2 { get; set; }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GetGridViewData(dataGridView1);
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            GetGridViewData(dataGridView1);
            this.Close();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmStaffSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            //this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int r = -1;
            if (txtsCode.Text != string.Empty) 
                r = GridViewFind(dataGridView1, cCode, txtsCode.Text);
            else if (txtsName.Text != string.Empty) 
                r = GridViewFind(dataGridView1, cName, txtsName.Text);

            if (r != -1)
            {
                dataGridView1.FirstDisplayedScrollingRowIndex = r;
                dataGridView1.CurrentCell = dataGridView1[cCode, r];
            }
        }

        /// <summary>
        /// コードまたは氏名で該当するデータの行を返します
        /// </summary>
        /// <param name="d">データグリッドビューオブジェクト</param>
        /// <param name="ColName">対象のカラム名</param>
        /// <param name="Val">検索する値</param>
        /// <returns>行のIndex</returns>
        private int GridViewFind(DataGridView d, string ColName, string Val)
        {
            int rtn = -1;

            for (int i = 0; i < d.Rows.Count; i++)
            {
                if (d[ColName, i].Value.ToString().Contains(Val))
                {
                    rtn = i;
                    break;
                }
            }

            return rtn;
        }

        private void txtsCode_TextChanged(object sender, EventArgs e)
        {
            if (txtsCode.Text.Length > 0) txtsName.Text = string.Empty;
        }

        private void txtsName_TextChanged(object sender, EventArgs e)
        {
            if (txtsName.Text.Length > 0) txtsCode.Text = string.Empty;
        }
    }
}
