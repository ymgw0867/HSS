using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HSS.Config;
using HSS.Model;
using HSS.ScanOcr;
using HSS.Combi;

namespace HSS
{
    public partial class Form1 : Form
    {
        //string sMstPath;
        //string pMstPath;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 解放する
            this.Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // フォーム最大値、最小値を設定する
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // ディレクトリ作成
            dirCreate(Properties.Settings.Default.DATAJ1);
            dirCreate(Properties.Settings.Default.DATAJ2);
            dirCreate(Properties.Settings.Default.DATAS1);
            dirCreate(Properties.Settings.Default.DATAS2);
            dirCreate(Properties.Settings.Default.instPath + Properties.Settings.Default.NG);
            dirCreate(Properties.Settings.Default.instPath + Properties.Settings.Default.OK);
            dirCreate(Properties.Settings.Default.instPath + Properties.Settings.Default.READ);

            // 環境設定テーブルに時間単位有休処理フィールドを追加
            alterTable();

            // 設定情報を取得します
            global.GetCommonYearMonth();
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     環境設定テーブルに時間単位有休処理フィールドを追加：2017/03/13 </summary>
        ///-------------------------------------------------------------------
        private void alterTable()
        {
            // データベースへ接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            try
            {
                // 環境設定テーブルに時間単位有休処理フィールドを追加：2017/03/13
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("alter table 環境設定 add column 時間単位有休処理 int NOT NULL");
                sCom.CommandText = sb.ToString();

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                // 時間単位有休処理フィールド登録済みのとき何もしない
            }
            finally
            {
                sCom.Connection.Close();
            }
        }

        /// <summary>
        /// マスター（ＣＳＶ）パスを取得します
        /// </summary>
        //private void GetStaffMstPath()
        //{
        //    Model.SysControl.SetDBConnect sDB = new Model.SysControl.SetDBConnect();
        //    OleDbCommand sCom = new OleDbCommand();
        //    sCom.Connection = sDB.cnOpen();

        //    string mySql = "select * from 環境設定";
        //    sCom.CommandText = mySql;
        //    OleDbDataReader dr = sCom.ExecuteReader();

        //    try
        //    {
        //        while (dr.Read())
        //        {
        //            sMstPath = dr["MSTS"].ToString();
        //            pMstPath = dr["MSTJ"].ToString();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "マスター（ＣＳＶ）パス取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        return;
        //    }
        //    finally
        //    {
        //        if (dr.IsClosed == false) dr.Close();
        //        if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
        //    }
        //}

        /// <summary>
        /// 該当パスが存在しなければ作成します
        /// </summary>
        /// <sPath>
        /// 該当パス
        /// </sPath>
        private void dirCreate(string sPath)
        {
            if (System.IO.Directory.Exists(sPath) == false)
                System.IO.Directory.CreateDirectory(sPath);
        }

        private void btnSetup_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmConfig frm = new frmConfig();
            frm.ShowDialog();
            this.Show();
        }

        private void btnPrePrint_Click(object sender, EventArgs e)
        {
            int _usrSel = global.END_SELECT;

            this.Hide();

            // 処理をするスタッフの種類を選択します
            PrePrint.frmUserSelect frm = new PrePrint.frmUserSelect();
            frm.ShowDialog();
            _usrSel = frm._usrSel;
            frm.Dispose();

            if (_usrSel == global.END_SELECT)
            {
                this.Show();
                return;
            }

            // プレ印刷画面表示
            PrePrint.frmPrePrint frmPre = new PrePrint.frmPrePrint(_usrSel);
            frmPre.ShowDialog();
            this.Show();
        }

        ///// <summary>
        ///// CSVデータをMDBへインサートする
        ///// </summary>
        //private void GetCsvToMdb(string csvPath, int usr)
        //{

        //    // CSVが存在しなければ終了する
        //    if (System.IO.File.Exists(csvPath) == false) return;

        //    bool md = false;

        //    // マスタレコード削除
        //    if (usr == global.STAFF_SELECT) md = StaffMstDelete();
        //    else if (usr == global.PART_SELECT) md = PartMstDelete();

        //    if (md == false) return;

        //    // CSV読み込み行カウント
        //    int csvLine = 0;

        //    // StreamReader の新しいインスタンスを生成する
        //    System.IO.StreamReader inFile = new System.IO.StreamReader(csvPath, Encoding.Default);

        //    // 読み込んだ結果をすべて格納するための変数を宣言する
        //    string stBuffer;

        //    // 読み込みできる文字がなくなるまで繰り返す
        //    while (inFile.Peek() >= 0)
        //    {
        //        // ファイルを 1 行ずつ読み込む
        //        stBuffer = inFile.ReadLine();
        //        csvLine++;

        //        // ヘッダの1行目は読み飛ばす
        //        if (csvLine > 1)
        //        {
        //            // カンマ区切りで分割して配列に格納する
        //            string[] stCSV = stBuffer.Split(',');

        //            //MDBへ登録する
        //            switch (usr)
        //            {
        //                case global.STAFF_SELECT:
        //                    ImportStaffCsv(stCSV);
        //                    break;
        //                case global.PART_SELECT:
        //                    ImportPartCsv(stCSV);
        //                    break;
        //                default:
        //                    break;
        //            }
        //        }
        //    }

        //    // StreamReader 閉じる
        //    inFile.Close();
        //}

        ///// <summary>
        ///// スタッフマスタの全レコードを削除します
        ///// </summary>
        ///// <returns>true:削除成功、false:削除失敗</returns>
        //private bool StaffMstDelete()
        //{
        //    bool let = false;

        //    // データベース接続文字列
        //    SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
        //    OleDbCommand sCom = new OleDbCommand();
        //    sCom.Connection = sDB.cnOpen();

        //    try
        //    {
        //        sCom.CommandText = "delete from スタッフマスタ ";
        //        sCom.ExecuteNonQuery();
        //        let = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "スタッフマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //    }
        //    finally
        //    {
        //        if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
        //    }

        //    return let;
        //}

        ///// <summary>
        ///// パートマスタの全レコードを削除します
        ///// </summary>
        ///// <returns>true:削除成功、false:削除失敗</returns>
        //private bool PartMstDelete()
        //{
        //    bool let = false;

        //    // データベース接続文字列
        //    SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
        //    OleDbCommand sCom = new OleDbCommand();
        //    sCom.Connection = sDB.cnOpen();

        //    try
        //    {
        //        sCom.CommandText = "delete from パートマスタ ";
        //        sCom.ExecuteNonQuery();
        //        let = true;
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "パートマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //    }
        //    finally
        //    {
        //        if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
        //    }

        //    return let;
        //}

        ///// <summary>
        ///// スタッフマスタインポート処理
        ///// </summary>
        ///// <param name="s">csv配列データ</param>
        //private void ImportStaffCsv(string[] s)
        //{
        //    //MDBへインポート
        //    // データベース接続文字列
        //    SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
        //    OleDbCommand sCom = new OleDbCommand();
        //    sCom.Connection = sDB.cnOpen();

        //    StringBuilder sb = new StringBuilder();

        //    try
        //    {
        //        sb.Clear();
        //        sb.Append("insert into スタッフマスタ (");
        //        sb.Append("オーダーコード,派遣先CD,派遣先名,派遣先部署,契約期間開始,契約期間終了,");
        //        sb.Append("スタッフコード,スタッフ名,開始時刻1,終了時刻1,契約内時間数1,給与区分,更新年月日) ");
        //        sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?)");

        //        sCom.CommandText = sb.ToString();
        //        sCom.Parameters.Clear();
        //        for (int i = 0; i < s.Length; i++)
        //        {
        //            // スタッフコードは7桁固定頭０埋めとします
        //            if (i == 6) sCom.Parameters.AddWithValue("@" + i.ToString(), s[i].PadLeft(7, '0'));
        //            else sCom.Parameters.AddWithValue("@" + i.ToString(), s[i]);
        //        }
        //        sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());
        //        sCom.ExecuteNonQuery();

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "スタッフマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        return;
        //    }
        //    finally
        //    {
        //        if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
        //    }
        //}

        ///// <summary>
        ///// パートマスタインポート処理
        ///// </summary>
        ///// <param name="s">csv配列データ</param>
        //private void ImportPartCsv(string[] s)
        //{
        //    //MDBへインポート
        //    // データベース接続文字列
        //    SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
        //    OleDbCommand sCom = new OleDbCommand();
        //    sCom.Connection = sDB.cnOpen();

        //    StringBuilder sb = new StringBuilder();

        //    try
        //    {
        //        sb.Clear();
        //        sb.Append("insert into パートマスタ (");
        //        sb.Append("勤務場所店番,勤務場所店名,姓,名,生年月日,個人番号,カナ氏名,母店番,母店名,");
        //        sb.Append("就業区分,当初雇用日,退職年月日,退職事由,雇用期間始,雇用期間終,");
        //        sb.Append("勤務時間始,勤務時間終,担務コード,担務内容,所定勤務日数,時給,副食費,");
        //        sb.Append("異動区分,更新年月日) ");
        //        sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");

        //        sCom.CommandText = sb.ToString();
        //        sCom.Parameters.Clear();
        //        for (int i = 0; i < s.Length; i++)
        //        {
        //            // 勤務場所店番は頭3桁を取得します
        //            if (i == 0) sCom.Parameters.AddWithValue("@" + i.ToString(), s[i].Substring(0,3));
        //            // 個人番号は7桁固定頭０埋めとします
        //            else if (i == 5) sCom.Parameters.AddWithValue("@" + i.ToString(), s[i].PadLeft(7, '0'));
        //            else sCom.Parameters.AddWithValue("@" + i.ToString(), s[i]);
        //        }
        //        sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());
        //        sCom.ExecuteNonQuery();

        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message, "パートマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        //        return;
        //    }
        //    finally
        //    {
        //        if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
        //    }
        //}

        private void btnOCR_Click(object sender, EventArgs e)
        {
            int _pcSel = global.END_SELECT;
            int _usrSel = 0;
            this.Hide();

            // PC、スタッフを選択する
            frmPcStaffSelect frmSel = new frmPcStaffSelect();
            frmSel.ShowDialog();
            _pcSel = frmSel._pcSel;
            _usrSel = frmSel._usrSel;

            frmSel.Dispose();

            if (_pcSel == global.END_SELECT)
            {
                this.Show();
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("スキャナで画像を読み取りOCR認識処理を行います。").Append(Environment.NewLine).Append(Environment.NewLine);
            sb.Append("よろしいですか？中止する場合は「いいえ」をクリックしてください。");
            if ((MessageBox.Show(sb.ToString(), this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.No))
            {
                this.Show();
                return;
            }
            else
            {
                // 勤務票スキャン、ＯＣＲ処理
                frmOCR frm = new frmOCR(_pcSel, _usrSel);
                frm.ShowDialog();
                this.Show();
            }
        }

        private void btnDataEntry_Click(object sender, EventArgs e)
        {
            // 環境設定年月の確認
            string msg = "設定年月は " + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月です。よろしいですか？";
            if (MessageBox.Show(msg, "勤務データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;
            
            int _usrSel = global.END_SELECT;

            this.Hide();

            // 処理をするスタッフの種類を選択します
            PrePrint.frmUserSelect frm = new PrePrint.frmUserSelect();
            frm.Text = "勤怠データ登録";
            frm.ShowDialog();
            _usrSel = frm._usrSel;
            frm.Dispose();

            if (_usrSel == global.END_SELECT)
            {
                this.Show();
                return;
            }

            // 勤怠データ登録画面表示
            ScanOcr.frmCorrect frmData = new frmCorrect(_usrSel, global.sEDITMODE);
            frmData.ShowDialog();
            this.Show();
        }

        private void btnDataCombi_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmCsvCombi frm = new frmCsvCombi();
            frm.ShowDialog();
            this.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            frmCalendar frm = new frmCalendar();
            frm.ShowDialog();
            this.Show();
        }

        private void btnDataAddNew_Click(object sender, EventArgs e)
        {
            int _usrSel = global.END_SELECT;

            this.Hide();

            // 処理をするスタッフの種類を選択します
            PrePrint.frmUserSelect frm = new PrePrint.frmUserSelect();
            frm.Text = "勤怠データ登録";
            frm.ShowDialog();
            _usrSel = frm._usrSel;
            frm.Dispose();

            if (_usrSel == global.END_SELECT)
            {
                this.Show();
                return;
            }

            // 勤怠データ登録画面表示
            ScanOcr.frmCorrect frmData = new frmCorrect(_usrSel,global.sADDMODE);
            frmData.ShowDialog();
            this.Show();
        }

    }
}
