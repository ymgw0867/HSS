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

namespace HSS.Config
{
    public partial class frmCalendar : Form
    {
        public frmCalendar()
        {
            InitializeComponent();
        }

        private void frmCalendar_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);
            GridViewSetting(dataGridView1);
            GridViewShow(dataGridView1);
            DispClr();
        }

        // 登録モード
        int _fMode = 0;

        // グリッドビューカラム名
        private string cYear = "c1";
        private string cMonth = "c2";
        private string cDay = "c3";
        private string cMemo = "c4";

        /// <summary>
        /// グリッドビューの定義を行います
        /// </summary>
        /// <param name="tempDGV">データグリッドビューオブジェクト</param>
        private void GridViewSetting(DataGridView tempDGV)
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
                tempDGV.Height = 382;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cYear, "年");
                tempDGV.Columns.Add(cMonth, "月");
                tempDGV.Columns.Add(cDay, "日");
                tempDGV.Columns.Add(cMemo, "名称");

                tempDGV.Columns[cYear].Width = 60;
                tempDGV.Columns[cMonth].Width = 40;
                tempDGV.Columns[cDay].Width = 40;
                tempDGV.Columns[cMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                tempDGV.Columns[cYear].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cMonth].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                tempDGV.Columns[cDay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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
        /// 休日データをグリッドビューへ表示します
        /// </summary>
        /// <param name="tempGrid">データグリッドビューオブジェクト</param>
        private void GridViewShow(DataGridView tempGrid)
        {
            //データベース接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            string mySql = "select * from 休日 order by 年,月,日";
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = mySql;
            dr = sCom.ExecuteReader();

            int iX = 0;
            tempGrid.RowCount = 0;

            while (dr.Read())
            {
                tempGrid.Rows.Add();

                tempGrid[cYear, iX].Value = dr["年"].ToString();
                tempGrid[cMonth, iX].Value = dr["月"].ToString();
                tempGrid[cDay, iX].Value = dr["日"].ToString();
                tempGrid[cMemo, iX].Value = dr["名称"].ToString();

                iX++;
            }

            dr.Close();
            sCom.Connection.Close();

            tempGrid.CurrentCell = null;
        }

        private void DispClr()
        {
            txtDate.Text = string.Empty;
            txtMemo.Text = string.Empty;

            btnUpdate.Enabled = false;
            btnDelete.Enabled = false;
            btnClr.Enabled = false;
            monthCalendar1.Enabled = true;
            _fMode = 0;
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            txtDate.Text = monthCalendar1.SelectionStart.ToString("yyyy/MM/dd (ddd)");
            btnUpdate.Enabled = true;
            btnClr.Enabled = true;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (txtDate.Text == string.Empty)
            {
                MessageBox.Show("日付が選択されていません", "休日設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            switch (_fMode)
            {
                case 0:
                    if (!dataSearch(monthCalendar1.SelectionStart))
                    {
                        if (MessageBox.Show(txtDate.Text + " を登録しますか？", "休日登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                        dataInsert(monthCalendar1.SelectionStart);
                    }
                    else
                    {
                        MessageBox.Show("既に登録済みの日付です", "休日設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    break;
                case 1:
                    if (MessageBox.Show(txtDate.Text + " を更新しますか？", "休日登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
                    dataUpdate(DateTime.Parse(txtDate.Text));
                    break;
                default:
                    break;
            }

            GridViewShow(dataGridView1);
            DispClr();
        }

        /// <summary>
        /// 休日テーブルに休日データを新規に登録する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        private void dataInsert(DateTime dt)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("insert into 休日 (年,月,日,名称,更新年月日) ");
            sb.Append("values (?,?,?,?,?) ");
            
            sCom.CommandText = sb.ToString();
            
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@year", dt.Year);
            sCom.Parameters.AddWithValue("@month", dt.Month);
            sCom.Parameters.AddWithValue("@day", dt.Day);
            sCom.Parameters.AddWithValue("@memo", txtMemo.Text);
            sCom.Parameters.AddWithValue("@date", DateTime.Today);
            
            sCom.ExecuteNonQuery();
            sCom.Connection.Close();
        }

        /// <summary>
        /// 休日データを更新する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        private void dataUpdate(DateTime dt)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("update 休日 set ");
            sb.Append("名称=?,更新年月日=? ");
            sb.Append("where 年=? and 月=? and 日=?");

            sCom.CommandText = sb.ToString();

            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@memo", txtMemo.Text);
            sCom.Parameters.AddWithValue("@date", DateTime.Today);
            sCom.Parameters.AddWithValue("@year", dt.Year);
            sCom.Parameters.AddWithValue("@month", dt.Month);
            sCom.Parameters.AddWithValue("@day", dt.Day);

            sCom.ExecuteNonQuery();
            sCom.Connection.Close();
        }

        /// <summary>
        /// 休日データを削除する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        private void dataDelete(DateTime dt)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("delete from 休日 ");
            sb.Append("where 年=? and 月=? and 日=?");

            sCom.CommandText = sb.ToString();

            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@year", dt.Year);
            sCom.Parameters.AddWithValue("@month", dt.Month);
            sCom.Parameters.AddWithValue("@day", dt.Day);

            sCom.ExecuteNonQuery();
            sCom.Connection.Close();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GetGridViewData(dataGridView1);
        }

        private void GetGridViewData(DataGridView g)
        {
            if (g.SelectedRows.Count == 0) return;

            int r = g.SelectedRows[0].Index;

            string y = g[cYear, r].Value.ToString();
            string m = g[cMonth, r].Value.ToString();
            string d = g[cDay, r].Value.ToString();

            txtDate.Text = DateTime.Parse(y + "/" + m + "/" + d).ToString("yyyy/MM/dd (ddd)");
            txtMemo.Text = g[cMemo, r].Value.ToString();

            btnUpdate.Enabled = true;
            btnDelete.Enabled = true;
            btnClr.Enabled = true;
            monthCalendar1.Enabled = false;
            _fMode = 1;
        }

        private void btnClr_Click(object sender, EventArgs e)
        {
            DispClr();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(txtDate.Text + " を削除してよろしいですか？", "休日削除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No) return;
            dataDelete(DateTime.Parse(txtDate.Text));
            DispClr();
            GridViewShow(dataGridView1);
        }

        /// <summary>
        /// 休日データを検索する
        /// </summary>
        /// <param name="dt">対象となる日付</param>
        /// <returns>true:データあり、false:データなし</returns>
        private bool dataSearch(DateTime dt)
        {
            bool _Rtn = false;

            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("select * from 休日 ");
            sb.Append("where 年=? and 月=? and 日=?");

            sCom.CommandText = sb.ToString();

            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@year", dt.Year);
            sCom.Parameters.AddWithValue("@month", dt.Month);
            sCom.Parameters.AddWithValue("@day", dt.Day);

            OleDbDataReader dR;
            dR = sCom.ExecuteReader();

            if (dR.HasRows) _Rtn = true;
            dR.Close();
            sCom.Connection.Close();

            return _Rtn;
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCalendar_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
