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
using Leadtools.Twain;

namespace HSS.Config
{
    public partial class frmConfig : Form
    {
        const int ADDNEW = 0;
        const int EDIT = 1;
        const string SEIREKI = "20";
        int fMode = ADDNEW;

        string sMstPath;
        string pMstPath;

        // TWAINでの取得用にTwainSessionを宣言します。
        TwainSession _twainSession = new TwainSession();

        public frmConfig()
        {
            InitializeComponent();
        }

        private void frmConfig_Load(object sender, EventArgs e)
        {
            // フォームの最大サイズ、最小サイズの設定
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // バックアップ自動削除月数コンボボックス値設定
            Utility.comboBkDel.Load(this.cmbBkdels);
            Utility.comboBkDel.Load(this.cmbBkdelp);

            // TwainSessionを初期化します
            _twainSession.Startup(this, "fkdl", "LEADTOOLS", "ver16.5J", "OCR", TwainStartupFlags.None);
            
            // データ表示
            DataShow();
        }

        private void DataShow()
        {
            // 表示項目クリア
            DispClear();

            // データベースの接続定義
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dr;

            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            OleDbConnection Cn = sDB.cnOpen();
            StringBuilder sb = new StringBuilder();

            sb.Clear();
            sb.Append("select * from 環境設定");
            sCom.CommandText = sb.ToString();
            sCom.Connection = Cn;
            dr = sCom.ExecuteReader();

            try
            {
                while (dr.Read())
                {
                    fMode = EDIT;
                    if (dr["SYEAR"].ToString().Length == 4)
                        txtYear.Text = dr["SYEAR"].ToString().Substring(2, 2);
                    else txtYear.Text = dr["SYEAR"].ToString();
                    txtMonth.Text = dr["SMONTH"].ToString();
                    //txtScanner.Text = dr["SCAN"].ToString();
                    txtMsts.Text = dr["MSTS"].ToString();
                    txtDat.Text = dr["DAT"].ToString();
                    txtTif.Text = dr["TIF"].ToString();
                    txtMstj.Text = dr["MSTJ"].ToString();
                    txtDat2.Text = dr["DAT2"].ToString();
                    txtTif2.Text = dr["TIF2"].ToString();

                    cmbBkdels.SelectedIndex = Utility.comboBkDel.selectedIndex(cmbBkdels, int.Parse(dr["BKDELS"].ToString()));
                    cmbBkdelp.SelectedIndex = Utility.comboBkDel.selectedIndex(cmbBkdelp, int.Parse(dr["BKDELP"].ToString()));

                    fMode = 1;

                    // 時間単位有休処理チェック：2017/03/13
                    if (Utility.StrToInt(Utility.NulltoStr(dr["時間単位有休処理"])) == 1)
                    {
                        chkTimeYukyu.CheckState = CheckState.Checked;
                    }
                    else
                    {
                        chkTimeYukyu.CheckState = CheckState.Unchecked;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (dr.IsClosed == false) dr.Close();
                if (Cn.State == ConnectionState.Open) Cn.Close();
            }
        }

        private void DispClear()
        {
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtScanner.Text = string.Empty;
            txtMsts.Text = string.Empty;
            txtTif.Text = string.Empty;
            txtDat.Text = string.Empty;
            txtMstj.Text = string.Empty;
            txtDat2.Text = string.Empty;
            txtTif2.Text = string.Empty;
            cmbBkdels.Text = string.Empty;
            cmbBkdelp.Text = string.Empty;

            if (TwainSession.IsAvailable(this))
            {
                btnScanner.Enabled = true;
                txtScanner.Enabled = true;
                txtScanner.Text = _twainSession.SelectedSourceName();
            }
            else
            {
                btnScanner.Enabled = false;
                txtScanner.Enabled = false;
            }
        }

        /// <summary>
        /// フォルダダイアログ選択
        /// </summary>
        /// <returns>フォルダー名</returns>
        private string userFolderSelect()
        {
            string fName = string.Empty;

            //出力フォルダの選択ダイアログの表示
            // FolderBrowserDialog の新しいインスタンスを生成する (デザイナから追加している場合は必要ない)
            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();

            // ダイアログの説明を設定する
            folderBrowserDialog1.Description = "フォルダを選択してください";

            // ルートになる特殊フォルダを設定する (初期値 SpecialFolder.Desktop)
            folderBrowserDialog1.RootFolder = System.Environment.SpecialFolder.Desktop;

            // 初期選択するパスを設定する
            folderBrowserDialog1.SelectedPath = @"C:\";

            // [新しいフォルダ] ボタンを表示する (初期値 true)
            folderBrowserDialog1.ShowNewFolderButton = true;

            // ダイアログを表示し、戻り値が [OK] の場合は、選択したディレクトリを表示する
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                fName = folderBrowserDialog1.SelectedPath + @"\";
            }
            else
            {
                // 不要になった時点で破棄する
                folderBrowserDialog1.Dispose();
                return fName;
            }

            // 不要になった時点で破棄する
            folderBrowserDialog1.Dispose();

            return fName;
        }

        private void btnDat_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            txtDat.Text = userFolderSelect();
        }

        private void btnTif_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            txtTif.Text = userFolderSelect();
        }

        private void btnDat2_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            txtDat2.Text = userFolderSelect();
        }

        private void BtnTif2_Click(object sender, EventArgs e)
        {
            //フォルダーを選択する
            txtTif2.Text = userFolderSelect();
        }

        private void btnMsts_Click(object sender, EventArgs e)
        {
            txtMsts.Text = getOpenFileName();
        }

        private string getOpenFileName()
        {
            string f = string.Empty;

            // ダイアログボックスの表示
            openFileDialog1.Title = "ファイルの選択";
            openFileDialog1.Filter = "データファイル(*.csv *.txt *.dat)|*.csv;*.txt;*.dat";
            openFileDialog1.InitialDirectory = Properties.Settings.Default.instPath;
            openFileDialog1.FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK) f = openFileDialog1.FileName;
            return f;
        }

        private void btnMstj_Click(object sender, EventArgs e)
        {
            txtMstj.Text = getOpenFileName();
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void txtYear_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();
            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;

            if (txtObj.Text == string.Empty) return;
            if (txtObj.Text.Length < 2) txtObj.Text = string.Format("{0:D2}", int.Parse(txtObj.Text)); 
        }

        private void txtYear_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = new TextBox();
            if (sender == txtYear) txtObj = txtYear;
            if (sender == txtMonth) txtObj = txtMonth;
            if (sender == txtScanner) txtObj = txtScanner;
            if (sender == txtMsts) txtObj = txtMsts;
            if (sender == txtDat) txtObj = txtDat;
            if (sender == txtTif) txtObj = txtTif;
            if (sender == txtMstj) txtObj = txtMstj;
            if (sender == txtDat2) txtObj = txtDat2;
            if (sender == txtTif2) txtObj = txtTif2;

            txtObj.SelectAll();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (errCheck())
            {
                switch (fMode)
                {
                    case ADDNEW:
                        AddNewRecord();
                        break;

                    case EDIT:
                        EditRecord();
                        break;
                }
            }

            // 設定対象年月を取得します
            global.GetCommonYearMonth();
            
            // マスター（ＣＳＶ）パスを取得します
            GetStaffMstPath();
            
            // スタッフCSVデータをMDBへインポートします
            GetCsvToMdb(sMstPath, Model.global.STAFF_SELECT);

            // パートCSVデータをMDBへインポートします
            GetCsvToMdb(pMstPath, Model.global.PART_SELECT);

            // 終了します
            this.Close();
        }

        /// <summary>
        /// 環境設定データ新規登録処理
        /// </summary>
        private void AddNewRecord()
        {
            OleDbCommand sCom = new OleDbCommand();
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            sCom.Connection = sDB.cnOpen();
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();
                sb.Append("insert into 環境設定 (");
                sb.Append("MSTJ,MSTS,TIF,TIF2,DAT,DAT2,BKDELS,BKDELP,PCADB,PCADB2,");
                sb.Append("SCAN,SYEAR,SMONTH,USRDATA,更新年月日,時間単位有休処理) values (");
                sb.Append("?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@msts", txtMstj.Text);
                sCom.Parameters.AddWithValue("@mstj", txtMsts.Text);
                sCom.Parameters.AddWithValue("@tif", txtTif.Text);
                sCom.Parameters.AddWithValue("@tif2", txtTif2.Text);
                sCom.Parameters.AddWithValue("@dat", txtDat.Text);
                sCom.Parameters.AddWithValue("@dat2", txtDat2.Text);
                sCom.Parameters.AddWithValue("@bkdels", int.Parse(cmbBkdels.Text));
                sCom.Parameters.AddWithValue("@bkdelp", int.Parse(cmbBkdelp.Text));
                sCom.Parameters.AddWithValue("@pcadb", string.Empty);
                sCom.Parameters.AddWithValue("@pcadb2", string.Empty);
                sCom.Parameters.AddWithValue("@scan", txtScanner.Text);
                sCom.Parameters.AddWithValue("@year", SEIREKI + txtYear.Text);
                sCom.Parameters.AddWithValue("@Month", txtMonth.Text);
                sCom.Parameters.AddWithValue("@userdata", string.Empty);
                sCom.Parameters.AddWithValue("@update", DateTime.Today.ToShortDateString());

                // 時間単位有休処理：2017/03/13
                if (chkTimeYukyu.CheckState == CheckState.Checked)
                {
                    sCom.Parameters.AddWithValue("@timeYukyu", 1);
                }
                else
                {
                    sCom.Parameters.AddWithValue("@timeYukyu", 0);
                }

                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 環境設定データ書き換え
        /// </summary>
        private void EditRecord()
        {
            OleDbCommand sCom = new OleDbCommand();
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            sCom.Connection = sDB.cnOpen();
            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();
                sb.Append("update 環境設定 set ");
                sb.Append("MSTJ=?,MSTS=?,TIF=?,TIF2=?,DAT=?,DAT2=?,BKDELS=?,BKDELP=?,");
                sb.Append("SCAN=?,SYEAR=?,SMONTH=?,更新年月日=?,時間単位有休処理=?");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@msts", txtMstj.Text);
                sCom.Parameters.AddWithValue("@mstj", txtMsts.Text);
                sCom.Parameters.AddWithValue("@tif", txtTif.Text);
                sCom.Parameters.AddWithValue("@tif2", txtTif2.Text);
                sCom.Parameters.AddWithValue("@dat", txtDat.Text);
                sCom.Parameters.AddWithValue("@dat2", txtDat2.Text);
                sCom.Parameters.AddWithValue("@bkdels", int.Parse(cmbBkdels.Text));
                sCom.Parameters.AddWithValue("@bkdelp", int.Parse(cmbBkdelp.Text));
                sCom.Parameters.AddWithValue("@scan", txtScanner.Text);
                sCom.Parameters.AddWithValue("@year", SEIREKI + txtYear.Text);
                sCom.Parameters.AddWithValue("@Month", txtMonth.Text);
                sCom.Parameters.AddWithValue("@update", DateTime.Today.ToShortDateString());

                // 時間単位有休処理：2017/03/13
                if (chkTimeYukyu.CheckState == CheckState.Checked)
                {
                    sCom.Parameters.AddWithValue("@timeYukyu", 1);
                }
                else
                {
                    sCom.Parameters.AddWithValue("@timeYukyu", 0);
                }
                
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private Boolean errCheck()
        {
            if (txtYear.Text == string.Empty)
            {
                MessageBox.Show("対象年が未入力です", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (txtMonth.Text == string.Empty)
            {
                MessageBox.Show("対象月が未入力です", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            int sMonth = int.Parse(txtMonth.Text);
            if (sMonth < 1 || sMonth > 12)
            {
                MessageBox.Show("対象月が正しくありません", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return false;
            }

            if (cmbBkdels.SelectedIndex == -1)
            {
                MessageBox.Show("スタッフのバックアップデータの自動削除設定を選択してください", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbBkdels.Focus();
                return false;
            }

            if (cmbBkdelp.SelectedIndex == -1)
            {
                MessageBox.Show("パートのバックアップデータの自動削除設定を選択してください", "環境設定", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbBkdelp.Focus();
                return false;
            }
            return true;
        }

        private void frmConfig_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnScanner_Click(object sender, EventArgs e)
        {
            try
            {
                _twainSession.SelectSource(string.Empty);
                txtScanner.Text = _twainSession.SelectedSourceName();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "スキャナの選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
        }

        /// <summary>
        /// マスター（ＣＳＶ）パスを取得します
        /// </summary>
        private void GetStaffMstPath()
        {
            Model.SysControl.SetDBConnect sDB = new Model.SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = sDB.cnOpen();

            string mySql = "select * from 環境設定";
            sCom.CommandText = mySql;
            OleDbDataReader dr = sCom.ExecuteReader();

            try
            {
                while (dr.Read())
                {
                    sMstPath = dr["MSTS"].ToString();
                    pMstPath = dr["MSTJ"].ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "マスター（ＣＳＶ）パス取得", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (dr.IsClosed == false) dr.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// CSVデータをMDBへインサートする
        /// </summary>
        private void GetCsvToMdb(string csvPath, int usr)
        {
            string hd = string.Empty;

            // データ件数
            int cTotal = 0;

            // CSVが存在しなければ終了する
            if (System.IO.File.Exists(csvPath) == false) return;

            bool md = false;

            // マスタレコード削除
            if (usr == global.STAFF_SELECT)
            {
                md = StaffMstDelete();
                hd = "スタッフ";
            }
            else if (usr == global.PART_SELECT)
            {
                md = PartMstDelete();
                hd = "パート";
            }

            if (md == false) return;

            // CSV読み込み行カウント
            int csvLine = 0;

            // StreamReader の新しいインスタンスを生成する
            System.IO.StreamReader inFile = new System.IO.StreamReader(csvPath, Encoding.Default);

            // 読み込んだ結果をすべて格納するための変数を宣言する
            string stBuffer;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // 件数を取得します
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                inFile.ReadLine();
                cTotal++;
            }

            // StreamReader の新しいインスタンスを生成する
            inFile = new System.IO.StreamReader(csvPath, Encoding.Default);

            // 読み込みできる文字がなくなるまで繰り返す
            while (inFile.Peek() >= 0)
            {
                // ファイルを 1 行ずつ読み込む
                stBuffer = inFile.ReadLine();
                csvLine++;

                //プログレスバー表示
                frmP.Text = hd + "ＣＳＶデータインポート実行中　" + csvLine.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = csvLine * 100 / cTotal;
                frmP.ProgressStep();

                // ヘッダの1行目は読み飛ばす
                if (csvLine > 1)
                {
                    // カンマ区切りで分割して配列に格納する
                    string[] stCSV = stBuffer.Split(',');

                    //MDBへ登録する
                    switch (usr)
                    {
                        case global.STAFF_SELECT:
                            ImportStaffCsv(stCSV);
                            break;
                        case global.PART_SELECT:
                            ImportPartCsv(stCSV);
                            break;
                        default:
                            break;
                    }
                }
            }

            // StreamReader 閉じる
            inFile.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;
        }

        /// <summary>
        /// スタッフマスタの全レコードを削除します
        /// </summary>
        /// <returns>true:削除成功、false:削除失敗</returns>
        private bool StaffMstDelete()
        {
            bool let = false;

            // データベース接続文字列
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = sDB.cnOpen();

            try
            {
                sCom.CommandText = "delete from スタッフマスタ ";
                sCom.ExecuteNonQuery();
                let = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "スタッフマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }

            return let;
        }

        /// <summary>
        /// パートマスタの全レコードを削除します
        /// </summary>
        /// <returns>true:削除成功、false:削除失敗</returns>
        private bool PartMstDelete()
        {
            bool let = false;

            // データベース接続文字列
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = sDB.cnOpen();

            try
            {
                sCom.CommandText = "delete from パートマスタ ";
                sCom.ExecuteNonQuery();
                let = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "パートマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }

            return let;
        }

        /// <summary>
        /// スタッフマスタインポート処理
        /// </summary>
        /// <param name="s">csv配列データ</param>
        private void ImportStaffCsv(string[] s)
        {
            //MDBへインポート
            // データベース接続文字列
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = sDB.cnOpen();

            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();
                sb.Append("insert into スタッフマスタ (");
                sb.Append("オーダーコード,派遣先CD,派遣先名,派遣先部署,契約期間開始,契約期間終了,");
                sb.Append("スタッフコード,スタッフ名,開始時刻1,終了時刻1,契約内時間数1,給与区分,更新年月日) ");
                sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                for (int i = 0; i < s.Length; i++)
                {
                    // スタッフコードは7桁固定頭０埋めとします
                    if (i == 6) sCom.Parameters.AddWithValue("@" + i.ToString(), s[i].PadLeft(7, '0'));
                    else sCom.Parameters.AddWithValue("@" + i.ToString(), s[i]);
                }
                sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());
                sCom.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "スタッフマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// パートマスタインポート処理
        /// </summary>
        /// <param name="s">csv配列データ</param>
        private void ImportPartCsv(string[] s)
        {
            //MDBへインポート
            // データベース接続文字列
            SysControl.SetDBConnect sDB = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = sDB.cnOpen();

            StringBuilder sb = new StringBuilder();

            try
            {
                sb.Clear();
                sb.Append("insert into パートマスタ (");
                sb.Append("勤務場所店番,勤務場所店名,姓,名,生年月日,個人番号,カナ氏名,母店番,母店名,");
                sb.Append("就業区分,当初雇用日,退職年月日,退職事由,雇用期間始,雇用期間終,");
                sb.Append("勤務時間始,勤務時間終,担務コード,担務内容,所定勤務日数,時給,副食費,");
                sb.Append("異動区分,更新年月日) ");
                sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                for (int i = 0; i < s.Length; i++)
                {
                    // 勤務場所店番は頭3桁を取得します
                    if (i == 0) sCom.Parameters.AddWithValue("@" + i.ToString(), s[i].Substring(0, 3));
                    // 個人番号は7桁固定頭０埋めとします
                    else if (i == 5) sCom.Parameters.AddWithValue("@" + i.ToString(), s[i].PadLeft(7, '0'));
                    else sCom.Parameters.AddWithValue("@" + i.ToString(), s[i]);
                }
                sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());
                sCom.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "パートマスターインポート", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }
    }
}
