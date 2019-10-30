using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using HSS.Model;
using Leadtools;
using Leadtools.Codecs;
using Leadtools.ImageProcessing;
using Leadtools.WinForms;

namespace HSS.ScanOcr
{
    public partial class frmCorrect : Form
    {
        public frmCorrect(int usrSel, int sMode)
        {
            InitializeComponent();

            // スタッフ指定
            _usrSel = usrSel;

            // 開始時処理モード
            _sMode = sMode;

            //tagを初期化
            this.Tag = END_MAKEDATA;

            // フォームキャプション
            SetformText();

            // 環境設定情報取得
            global.GetCommonYearMonth();

            // 登録モードのとき新規データを追加します
            if (sMode == global.sADDMODE) AddNewData(_usrSel);
        }

        //終了ステータス
        const string END_BUTTON = "btn";
        const string END_MAKEDATA = "data";
        const string END_CONTOROL = "close";

        // 指定スタッフ
        private int _usrSel;

        // 開始時処理モード
        private int _sMode;

        // 入力パス
        private string _InPath;

        // データグリッドビューカラム定義
        private string cDay = "col1";
        private string cWeek = "col2";
        private string cMark = "col3";
        private string cSH = "col4";
        private string cS = "col5";
        private string cSM = "col6";
        private string cEH = "col7";
        private string cE = "col8";
        private string cEM = "col9";
        private string cKyuka = "col10";
        private string cKyukei = "col11";
        private string cHiru1 = "col12";
        private string cHiru2 = "col13";
        private string cTeisei = "col14";
        private string cID = "col15";

        //datagridview表示行数
        private const int _MULTIGYO = 31;

        // MDBデータキー配列
        Entity.kData[] inData;

        //カレントデータインデックス
        private int _cI;

        /// <summary>
        /// データグリッドビューの定義を行います
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
                tempDGV.DefaultCellStyle.Font = new Font("メイリオ", 10, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                tempDGV.ColumnHeadersHeight = 22;
                tempDGV.RowTemplate.Height = 22;

                // 全体の高さ
                tempDGV.Height = 706;

                // 全体の幅
                //tempDGV.Width = 583;

                // 奇数行の色
                //tempDGV.AlternatingRowsDefaultCellStyle.BackColor = Color.LightBlue;

                //各列幅指定
                tempDGV.Columns.Add(cDay, "日");
                tempDGV.Columns.Add(cWeek, "曜");
                tempDGV.Columns.Add(cMark, "○");

                ////DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                ////tempDGV.Columns.Add(column);
                ////tempDGV.Columns[2].Name = cMark;
                ////tempDGV.Columns[2].HeaderText = "○";

                tempDGV.Columns.Add(cSH, "開");
                tempDGV.Columns.Add(cS, string.Empty);
                tempDGV.Columns.Add(cSM, "始");
                tempDGV.Columns.Add(cEH, "終");
                tempDGV.Columns.Add(cE, string.Empty);
                tempDGV.Columns.Add(cEM, "了");
                tempDGV.Columns.Add(cKyuka, "休暇");
                tempDGV.Columns.Add(cKyukei, "休憩");
                tempDGV.Columns.Add(cHiru1, "昼１");
                tempDGV.Columns.Add(cHiru2, "昼２");

                DataGridViewCheckBoxColumn column = new DataGridViewCheckBoxColumn();
                tempDGV.Columns.Add(column);
                tempDGV.Columns[13].Name = cTeisei;
                tempDGV.Columns[13].HeaderText = "訂正";
                tempDGV.Columns.Add(cID, "ID");
                tempDGV.Columns[cID].Visible = false; // IDカラムは非表示とする

                // 各列の定義を行う
                foreach (DataGridViewColumn c in tempDGV.Columns)
                {
                    // 幅
                    c.Width = 40;

                    // 表示位置、編集可否
                    if (c.Name == cDay)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cWeek)
                    {
                        c.Width = 28;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cMark)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cSH)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cS)
                    {
                        c.Width = 16;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cSM)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cEH)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cE)
                    {
                        c.Width = 16;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = true;
                    }
                    if (c.Name == cEM)
                    {
                        c.Width = 30;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cKyuka)
                    {
                        c.Width = 40;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cKyukei)
                    {
                        c.Width = 40;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cHiru1)
                    {
                        c.Width = 40;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cHiru2)
                    {
                        c.Width = 40;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cTeisei)
                    {
                        //c.Width = 40;
                        c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                        c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                        c.ReadOnly = false;
                    }
                    if (c.Name == cID)
                    {
                        c.ReadOnly = true;
                    }

                    // 入力可能桁数
                    if (c.Name != cTeisei)
                    {
                        DataGridViewTextBoxColumn col = (DataGridViewTextBoxColumn)c;

                        if (c.Name == cMark) col.MaxInputLength = 1;
                        if (c.Name == cSH) col.MaxInputLength = 2;
                        if (c.Name == cSM) col.MaxInputLength = 2;
                        if (c.Name == cEH) col.MaxInputLength = 2;
                        if (c.Name == cEM) col.MaxInputLength = 2;
                        if (c.Name == cKyuka) col.MaxInputLength = 1;
                        if (c.Name == cKyukei) col.MaxInputLength = 1;
                        if (c.Name == cHiru1) col.MaxInputLength = 1;
                        if (c.Name == cHiru2) col.MaxInputLength = 1;
                    }
                }

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.CellSelect;
                tempDGV.MultiSelect = false;

                // 編集可とする
                //tempDGV.ReadOnly = false;

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

                // 編集モード
                tempDGV.EditMode = DataGridViewEditMode.EditOnEnter;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
            }
        }

        private void SetformText()
        {
            if (_usrSel == global.STAFF_SELECT)
                this.Text = "勤怠データ スタッフ ― 登録  処理年月：";
            else if (_usrSel == global.PART_SELECT)
                this.Text = "勤怠データ パートタイマー ― 登録  処理年月：";

            this.Text += string.Format("{0}年{1}月", global.sYear.ToString(), global.sMonth.ToString());
        }

        private void frmCorrect_Load(object sender, EventArgs e)
        {
            // フォーム最大値
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            // フォーム最小値
            Utility.WindowsMinSize(this, this.Width, this.Height);
            // グリッド定義
            GridViewSetting(dg1);

            // スタッフ・パートによる表示・非表示切り替え
            if (_usrSel == global.STAFF_SELECT)
            {
                this.txtShozokuCode.MaxLength = 8;
                this.label7.Enabled = true;
                this.txtOrderCode.Enabled = true;

                if (global.sTimeYukyu == 0)
                {
                    this.label10.Enabled = false;
                    this.txtNichu.Enabled = false;
                    this.label11.Enabled = false;

                    this.label13.Enabled = false;
                    this.label12.Enabled = false;
                    this.txtChisou.Enabled = false;
                }
                else
                {
                    this.label10.Enabled = true;
                    this.txtNichu.Enabled = true;
                    this.label11.Enabled = true;

                    this.label13.Enabled = true;
                    this.label12.Enabled = true;
                    this.txtChisou.Enabled = true;
                }

                this.label16.Enabled = false;
                this.txtChuTl.Enabled = false;
                this.label17.Enabled = false;
                //this.panel3.Visible = false;
                this.panel3.Visible = true;     // 2017/04/24

                this.dg1.Columns[cKyuka].Width = 70;
                this.dg1.Columns[cKyukei].Width = 70;
                this.dg1.Columns[cKyukei].HeaderText = "休憩なし";
                this.dg1.Columns[cHiru1].ReadOnly = false;
                this.dg1.Columns[cHiru2].ReadOnly = false;
                this.dg1.Columns[cHiru1].Visible = false;
                this.dg1.Columns[cHiru2].Visible = false;
                this.dg1.Columns[cHiru1].HeaderText = "-";
                this.dg1.Columns[cHiru2].HeaderText = "-";
            }
            else if (_usrSel == global.PART_SELECT)
            {
                this.txtShozokuCode.MaxLength = 3;
                this.label7.Enabled = false;
                this.txtOrderCode.Enabled = false;
                this.label10.Enabled = true;
                this.txtNichu.Enabled = true;
                this.label11.Enabled = true;
                this.label13.Enabled = true;
                this.label12.Enabled = true;
                this.txtChisou.Enabled = true;
                this.label16.Enabled = true;
                this.txtChuTl.Enabled = true;
                this.label17.Enabled = true;
                this.panel3.Visible = true;

                this.dg1.Columns[cKyukei].Width = 40;
                this.dg1.Columns[cKyukei].HeaderText = "休憩";
                this.dg1.Columns[cHiru1].ReadOnly = false;
                this.dg1.Columns[cHiru2].ReadOnly = false;
                this.dg1.Columns[cHiru1].Visible = true;
                this.dg1.Columns[cHiru2].Visible = true;
                this.dg1.Columns[cHiru1].HeaderText = "昼①";
                this.dg1.Columns[cHiru2].HeaderText = "昼②";
            }

            // 入力ファイルフォルダ取得
            _InPath = exisInFileDir();

            // CSVデータMDBへ登録
            GetCsvDataToMDB();

            //MDB件数カウント
            if (CountMDB(_usrSel) == 0)
            {
                MessageBox.Show("対象となる勤務票データがありません", "勤務票データ登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //終了処理
                this.Close();
                return;
            }

            // MDBへ所定勤務情報をセットします
            GetMstData();

            // MDBデータキー項目読み込み
            inData = LoadMdbID(_usrSel);

            //エラー情報初期化
            ErrInitial();

            // 開始時処理モードによって表示するデータをコントロールします
            switch (_sMode)
            {
                // 登録モードのとき最後のレコードを表示
                case global.sADDMODE:
                    _cI = inData.Length - 1;
                    break;

                // 編集モードのとき先頭レコードを表示
                case global.sEDITMODE:
                    _cI = 0;
                    break;

                default:
                    break;
            }

            // データ表示
            DataShow(_cI, inData, this.dg1);
        }

        /// <summary>
        /// 勤務票ヘッダレコードに所定勤務時間をセットします
        /// </summary>
        private void GetMstData()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            OleDbDataReader dR;

            StringBuilder sb = new StringBuilder();

            // 勤務票ヘッダ情報から個人番号を取得します
            sCom.Parameters.Clear();
            if (_usrSel == global.STAFF_SELECT)
            {
                sb.Clear();
                sb.Append("select 勤務票ヘッダ.個人番号,スタッフマスタ.開始時刻1 as 開始時刻,スタッフマスタ.終了時刻1 as 終了時刻 ");
                sb.Append("from 勤務票ヘッダ inner join スタッフマスタ ");
                sb.Append("on 勤務票ヘッダ.個人番号 = スタッフマスタ.スタッフコード ");
                sCom.CommandText = sb.ToString();
            }
            else if (_usrSel == global.PART_SELECT)
            {
                sb.Clear();
                sb.Append("select 勤務票ヘッダ.個人番号,パートマスタ.勤務時間始 as 開始時刻,パートマスタ.勤務時間終 as 終了時刻 ");
                sb.Append("from 勤務票ヘッダ inner join パートマスタ ");
                sb.Append("on 勤務票ヘッダ.個人番号 = パートマスタ.個人番号 ");
                sCom.CommandText = sb.ToString();
            }
            dR = sCom.ExecuteReader();

            OleDbCommand sCom2 = new OleDbCommand();
            sCom2.Connection = Con.cnOpen();
            sb.Clear();
            sb.Append("update 勤務票ヘッダ set ");
            sb.Append("所定開始時間=?,所定終了時間=?,午前半休開始時間=?,午前半休終了時間=?, ");
            sb.Append("午後半休開始時間=?,午後半休終了時間=?,所定勤務時間=? ");
            sb.Append("where 個人番号=?");
            sCom2.CommandText = sb.ToString();

            string TimeS = string.Empty;
            string TimeE = string.Empty;

            while (dR.Read())
            {
                sCom2.Parameters.Clear();

                // 所定勤務開始、終了時間
                if (_usrSel == global.PART_SELECT)
                {
                    DateTime ds = DateTime.FromOADate(double.Parse(dR["開始時刻"].ToString()));
                    DateTime de = DateTime.FromOADate(double.Parse(dR["終了時刻"].ToString()));
                    sCom2.Parameters.AddWithValue("@st", ds.ToShortTimeString().PadLeft(5, '0'));
                    sCom2.Parameters.AddWithValue("@et", de.ToShortTimeString().PadLeft(5, '0'));
                }
                else if (_usrSel == global.STAFF_SELECT)
                {
                    sCom2.Parameters.AddWithValue("@st", dR["開始時刻"].ToString().PadLeft(5, '0'));
                    sCom2.Parameters.AddWithValue("@et", dR["終了時刻"].ToString().PadLeft(5, '0'));
                }

                // 開始時間取得
                if (_usrSel == global.PART_SELECT)
                {
                    DateTime ds = DateTime.FromOADate(double.Parse(dR["開始時刻"].ToString()));
                    DateTime de = DateTime.FromOADate(double.Parse(dR["終了時刻"].ToString()));
                    TimeS = ds.ToShortTimeString().PadLeft(5, '0');
                    TimeE = de.ToShortTimeString().PadLeft(5, '0');
                }
                else if (_usrSel == global.STAFF_SELECT)
                {
                    TimeS = dR["開始時刻"].ToString().PadLeft(5, '0');
                    TimeE = dR["終了時刻"].ToString().PadLeft(5, '0');
                }

                int TimeJ = (Utility.TimeToMin(TimeE) - Utility.TimeToMin(TimeS) - 60) / 2;
                DateTime dd = DateTime.Parse(TimeS);

                // 午前半休終了時間取得 
                string amE = dd.AddMinutes(TimeJ).ToShortTimeString();

                // 午後半休開始時間取得 
                string pmS = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                // 所定勤務時間 2012/04/21
                //double ShoteiTime = double.Parse(TimeJ.ToString()) / 60;
                double ShoteiTime = Utility.fncTime10(double.Parse(TimeJ.ToString()));

                sCom2.Parameters.AddWithValue("@amS", TimeS);
                sCom2.Parameters.AddWithValue("@amE", amE);
                sCom2.Parameters.AddWithValue("@pmS", pmS);
                sCom2.Parameters.AddWithValue("@pmE", TimeE);
                sCom2.Parameters.AddWithValue("@Sho", ShoteiTime);

                sCom2.Parameters.AddWithValue("@sCode", dR["個人番号"].ToString());

                sCom2.ExecuteNonQuery();
            }
            dR.Close();
            sCom.Connection.Close();
            sCom2.Connection.Close();
        }

        /// <summary>
        /// 個人別の所定勤務時間を取得します
        /// </summary>
        /// <param name="sCode"></param>
        private void GetShoteiData(string sCode)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            OleDbDataReader dR;

            StringBuilder sb = new StringBuilder();
            sCom.Parameters.Clear();

            // スタッフマスタから個人番号を取得します
            if (_usrSel == global.STAFF_SELECT)
            {
                sb.Clear();
                sb.Append("select 派遣先CD, 派遣先名, 派遣先部署, スタッフ名, 開始時刻1 as 開始時刻,終了時刻1 as 終了時刻,オーダーコード ");
                sb.Append("from スタッフマスタ ");
                sb.Append("where スタッフコード = ?");
            }
            // パートマスタから個人番号を取得します
            else if (_usrSel == global.PART_SELECT)
            {
                sb.Clear();
                sb.Append("select 勤務場所店番, 勤務場所店名, 姓, 名, 勤務時間始 as 開始時刻,勤務時間終 as 終了時刻 ");
                sb.Append("from パートマスタ ");
                sb.Append("where 個人番号 = ?");
            }

            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@Code",sCode);
            dR = sCom.ExecuteReader();

            string TimeS = string.Empty;
            string TimeE = string.Empty;

            // 該当項目を初期化します
            this.lblName.Text = string.Empty;
            this.txtShozokuCode.Text = string.Empty;
            this.txtOrderCode.Text = string.Empty;
            this.lblShozoku.Text = string.Empty;
            this.lblShoteiS.Text = string.Empty;
            this.lblShoteiE.Text = string.Empty;
            this.lblHanAmS.Text = string.Empty;
            this.lblHanAmE.Text = string.Empty;
            this.lblHanPmS.Text = string.Empty;
            this.lblHanPmE.Text = string.Empty;

            // マスター情報を表示します
            while (dR.Read())
            {
                // 所定勤務開始、終了時間
                if (_usrSel == global.PART_SELECT)
                {
                    DateTime ds = DateTime.FromOADate(double.Parse(dR["開始時刻"].ToString()));
                    DateTime de = DateTime.FromOADate(double.Parse(dR["終了時刻"].ToString()));
                    this.lblShoteiS.Text = ds.ToShortTimeString().PadLeft(5, '0');
                    this.lblShoteiE.Text = de.ToShortTimeString().PadLeft(5, '0');
                    TimeS = ds.ToShortTimeString().PadLeft(5, '0');
                    TimeE = de.ToShortTimeString().PadLeft(5, '0');

                    this.lblName.Text = dR["姓"].ToString() + " " + dR["名"].ToString();
                    this.txtShozokuCode.Text = dR["勤務場所店番"].ToString();
                    this.lblShozoku.Text = dR["勤務場所店名"].ToString();
                }
                else if (_usrSel == global.STAFF_SELECT)
                {
                    this.lblShoteiS.Text = dR["開始時刻"].ToString().PadLeft(5, '0');
                    this.lblShoteiE.Text = dR["終了時刻"].ToString().PadLeft(5, '0');
                    TimeS = dR["開始時刻"].ToString().PadLeft(5, '0');
                    TimeE = dR["終了時刻"].ToString().PadLeft(5, '0');

                    this.lblName.Text = dR["スタッフ名"].ToString();
                    this.txtShozokuCode.Text = dR["派遣先CD"].ToString().Substring(1, 8);

                    if (int.Parse(dR["派遣先CD"].ToString().Substring(0, 1)) >= 2)
                        this.lblShozoku.Text = dR["派遣先名"].ToString() + " " + dR["派遣先部署"].ToString();
                    else this.lblShozoku.Text = dR["派遣先部署"].ToString();
                    this.txtOrderCode.Text = dR["オーダーコード"].ToString();
                }

                int TimeJ = (Utility.TimeToMin(TimeE) - Utility.TimeToMin(TimeS) - 60) / 2;
                DateTime dd = DateTime.Parse(TimeS);

                // 午前半休終了時間取得 
                string amE = dd.AddMinutes(TimeJ).ToShortTimeString();

                // 午後半休開始時間取得 
                string pmS = dd.AddMinutes(TimeJ + 60).ToShortTimeString();

                // 所定勤務時間
                //double ShoteiTime = double.Parse(TimeJ.ToString()) / 60;
                double ShoteiTime = Utility.fncTime10(double.Parse(TimeJ.ToString()));

                lblHanAmS.Text = TimeS;
                lblHanAmE.Text = amE;
                lblHanPmS.Text = pmS;
                lblHanPmE.Text = TimeE;
                inData[_cI]._ShoTime = ShoteiTime;
            }

            dR.Close();
            sCom.Connection.Close();
        }

        /// <summary>
        /// 入力対象データファイルパスを取得します
        /// </summary>
        private string exisInFileDir()
        {
            string outPath = string.Empty;
            string formatFileName = string.Empty;

            // ＯＣＲ出力先パス

            switch (_usrSel)
            {
                case global.STAFF_SELECT:
                    if (Properties.Settings.Default.PC == 1)
                        outPath = Properties.Settings.Default.DATAS1;
                    else outPath = Properties.Settings.Default.DATAS2;
                    break;

                case global.PART_SELECT:
                    if (Properties.Settings.Default.PC == 1)
                        outPath = Properties.Settings.Default.DATAJ1;
                    else outPath = Properties.Settings.Default.DATAJ2;
                    break;
            }
            return outPath;
        }

        private void dg1_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == 3 || e.ColumnIndex == 4 || e.ColumnIndex == 6 || e.ColumnIndex == 7)
            {
                e.AdvancedBorderStyle.Right = DataGridViewAdvancedCellBorderStyle.None;

                if (e.ColumnIndex == 4 || e.ColumnIndex == 5 || e.ColumnIndex == 7 || e.ColumnIndex == 8)
                {
                    e.AdvancedBorderStyle.Left = DataGridViewAdvancedCellBorderStyle.None;
                }
                //else
                //    e.AdvancedBorderStyle.Left = dg1.AdvancedCellBorderStyle.Left;
            }
        }

        ///---------------------------------------------
        /// <summary>
        ///     CSVデータをMDBへインサートする </summary>
        ///---------------------------------------------
        private void GetCsvDataToMDB()
        {
            //CSVファイル数をカウント
            string[] inCsv = System.IO.Directory.GetFiles(_InPath, "*.csv");

            //CSVファイルがなければ終了
            int cTotal = 0;
            if (inCsv.Length == 0) return;
            else cTotal = inCsv.Length;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            // データベースへ接続
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                //CSVデータをMDBへ取込
                int cCnt = 0;
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    //件数カウント
                    cCnt++;

                    //プログレスバー表示
                    frmP.Text = "OCR変換CSVデータロード中　" + cCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = cCnt / cTotal * 100;
                    frmP.ProgressStep();

                    ////////OCR処理対象のCSVファイルかファイル名の文字数を検証する
                    //////string fn = Path.GetFileName(files);

                    // CSVファイルインポート
                    var s = File.ReadAllLines(files, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // MDBへ登録する
                        // 勤務記録ヘッダテーブル
                        StringBuilder sb = new StringBuilder();
                        sb.Clear();
                        sb.Append("insert into 勤務票ヘッダ ");
                        sb.Append("(ID,シートID,個人番号,年,月,勤務場所コード,オーダーコード,離席時間1,");
                        sb.Append("離席時間2,日中,遅早,職場離脱,出勤日数合計,昼食回数,画像名,更新年月日,データ区分) ");
                        sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");

                        sCom.CommandText = sb.ToString();
                        sCom.Parameters.Clear();

                        switch (_usrSel)
                        {
                            case global.STAFF_SELECT:

                                // ID
                                sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[0], 17));
                                // シートＩＤ
                                sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[2], 1));
                                // 個人番号
                                sCom.Parameters.AddWithValue("@kjn", Utility.GetStringSubMax(stCSV[3], 7));
                                // 年
                                sCom.Parameters.AddWithValue("@year", Utility.GetStringSubMax(stCSV[4], 4));
                                // 月
                                sCom.Parameters.AddWithValue("@month", Utility.GetStringSubMax(stCSV[5], 2));
                                // 所属コード
                                sCom.Parameters.AddWithValue("@Szk", Utility.GetStringSubMax(stCSV[6], 8));
                                // オーダーコード
                                sCom.Parameters.AddWithValue("@order", Utility.GetStringSubMax(stCSV[7] + stCSV[8], 7));
                                // 離席時間
                                // 「無記入で反応」対策 2017/05/22
                                string riseki = Utility.GetStringSubMax(stCSV[9], 3).Trim();
                                if (riseki == "5")
                                {
                                    sCom.Parameters.AddWithValue("@riseki1", string.Empty);
                                }
                                else
                                {
                                    sCom.Parameters.AddWithValue("@riseki1", riseki);
                                }

                                sCom.Parameters.AddWithValue("@riseki2", Utility.GetStringSubMax(stCSV[10], 1));

                                // 時間単位有休処理を行うときのみ日中・遅早を読み込む 2017/03/13
                                if (global.sTimeYukyu == 0)
                                {
                                    // 日中
                                    sCom.Parameters.AddWithValue("@nichu", string.Empty);
                                    // 遅早
                                    sCom.Parameters.AddWithValue("@Chisou", string.Empty);
                                }
                                else
                                {
                                    // 日中
                                    sCom.Parameters.AddWithValue("@nichu", Utility.GetStringSubMax(stCSV[13], 2));  // 2017/03/13
                                    // 遅早
                                    sCom.Parameters.AddWithValue("@Chisou", Utility.GetStringSubMax(stCSV[14], 2)); // 2017/03/13
                                }

                                // 職場離脱
                                sCom.Parameters.AddWithValue("@Shokuba", stCSV[11]);
                                // 出勤日数合計
                                sCom.Parameters.AddWithValue("@ShukinTL", Utility.GetStringSubMax(stCSV[12], 2));
                                // 昼食回数合計
                                sCom.Parameters.AddWithValue("@ChushokuTL", string.Empty);
                                // 画像名
                                sCom.Parameters.AddWithValue("@IMG", Utility.GetStringSubMax(stCSV[1], 21));
                                // 更新年月日
                                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());
                                // データ区分
                                sCom.Parameters.AddWithValue("@dKbn", _usrSel);

                                break;

                            case global.PART_SELECT:

                                // ID
                                sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[0], 17));
                                // シートＩＤ
                                sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[2], 1));
                                // 個人番号
                                sCom.Parameters.AddWithValue("@kjn", Utility.GetStringSubMax(stCSV[3], 7));
                                // 年
                                sCom.Parameters.AddWithValue("@year", Utility.GetStringSubMax(stCSV[4], 4));
                                // 月
                                sCom.Parameters.AddWithValue("@month", Utility.GetStringSubMax(stCSV[5], 2));
                                // 所属コード
                                sCom.Parameters.AddWithValue("@Szk", Utility.GetStringSubMax(stCSV[6], 3));
                                // オーダーコード
                                sCom.Parameters.AddWithValue("@order", string.Empty);
                                // 離席時間
                                sCom.Parameters.AddWithValue("@riseki1", Utility.GetStringSubMax(stCSV[7], 3));
                                sCom.Parameters.AddWithValue("@riseki2", Utility.GetStringSubMax(stCSV[8], 1));
                                // 日中
                                sCom.Parameters.AddWithValue("@nichu", Utility.GetStringSubMax(stCSV[9], 2));
                                // 遅早
                                sCom.Parameters.AddWithValue("@Chisou", Utility.GetStringSubMax(stCSV[10], 2));
                                // 職場離脱
                                sCom.Parameters.AddWithValue("@Shokuba", Utility.GetStringSubMax(stCSV[11], 1));
                                // 出勤日数合計
                                sCom.Parameters.AddWithValue("@ShukinTL", Utility.GetStringSubMax(stCSV[12], 2));
                                // 昼食回数合計
                                sCom.Parameters.AddWithValue("@ChushokuTL", Utility.GetStringSubMax(stCSV[13], 2));
                                // 画像名
                                sCom.Parameters.AddWithValue("@IMG", Utility.GetStringSubMax(stCSV[1], 21));
                                // 更新年月日
                                sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());
                                // データ区分
                                sCom.Parameters.AddWithValue("@dKbn", _usrSel);
                                break;
                        }

                        // テーブル書き込み
                        sCom.ExecuteNonQuery();

                        // 勤務票明細テーブル
                        sb.Clear();
                        sb.Append("insert into 勤務票明細 ");
                        sb.Append("(ヘッダID,日付,マーク,開始時,開始分,終了時,終了分,休暇,休憩なし,昼1,昼2,訂正,更新年月日,休日) ");
                        sb.Append("values (?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                        sCom.CommandText = sb.ToString();

                        int sDays = 0;
                        DateTime dt;

                        switch (_usrSel)
                        {
                            case global.STAFF_SELECT:
                                for (int i = 15; i <= 255; i += 8)
                                {
                                    sCom.Parameters.Clear();
                                    sDays++;

                                    // 存在する日付のときにMDBへ登録する
                                    string tempDt = global.sYear + "/" + global.sMonth + "/" + sDays.ToString();
                                    if (DateTime.TryParse(tempDt, out dt))
                                    {
                                        // ヘッダID
                                        sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[0], 17));

                                        // 日付
                                        sCom.Parameters.AddWithValue("@Days", sDays);

                                        // マーク
                                        sCom.Parameters.AddWithValue("@yukyu", stCSV[i]);

                                        // 開始時間                                        
                                        string hh = Utility.GetStringSubMax(stCSV[i + 1], 2);
                                        string mm = Utility.GetStringSubMax(stCSV[i + 2], 2);

                                        // 終了時間
                                        string hh2 = Utility.GetStringSubMax(stCSV[i + 3], 2);
                                        string mm2 = Utility.GetStringSubMax(stCSV[i + 4], 2);

                                        // 「無記入で反応」対策 2017/04/28
                                        if (hh.Trim() == string.Empty || mm.Trim() == string.Empty)
                                        {
                                            hh = string.Empty;
                                            mm = string.Empty;
                                        }
                                        
                                        // 「無記入で反応」対策 2017/04/28
                                        if (hh2.Trim() == string.Empty || mm2.Trim() == string.Empty)
                                        {
                                            hh2 = string.Empty;
                                            mm2 = string.Empty;
                                        }

                                        // 「無記入で反応」対策 2017/05/22
                                        // 開始、終了いずれかが無記入のとき両方無記入とする
                                        if ((hh + mm) == string.Empty || (hh2 + mm2) == string.Empty)
                                        {
                                            hh = string.Empty;
                                            mm = string.Empty;
                                            hh2 = string.Empty;
                                            mm2 = string.Empty;
                                        }

                                        sCom.Parameters.AddWithValue("@sh", hh);
                                        sCom.Parameters.AddWithValue("@sm", mm);
                                        
                                        sCom.Parameters.AddWithValue("@eh", hh2);
                                        sCom.Parameters.AddWithValue("@em", mm2);
                                        
                                        // 休暇
                                        sCom.Parameters.AddWithValue("@kyukei", stCSV[i + 5]);

                                        // 休憩なし
                                        // 「無記入で反応」対策：開始終了が無記入なら休憩なしも無記入とする 2017/05/22
                                        if (hh == string.Empty && mm == string.Empty &&
                                            hh2 == string.Empty && mm2 == string.Empty)
                                        {
                                            sCom.Parameters.AddWithValue("@kyunashi", global.FLGOFF);
                                        }
                                        else
                                        {
                                            sCom.Parameters.AddWithValue("@kyunashi", stCSV[i + 6]);
                                        }

                                        // 昼1
                                        sCom.Parameters.AddWithValue("@hiru1", 0);

                                        // 昼2
                                        sCom.Parameters.AddWithValue("@hiru2", 0);

                                        // 訂正
                                        sCom.Parameters.AddWithValue("@teisei", stCSV[i + 7]);

                                        // 更新年月日
                                        sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());

                                        // 休日区分の設定
                                        // ①曜日で判断します
                                        string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                                        int sHol = 0;
                                        if (Youbi == "土")
                                        {
                                            sHol = global.hSATURDAY;
                                        }
                                        else if (Youbi == "日")
                                        {
                                            sHol = global.hHOLIDAY;
                                        }
                                        else
                                        {
                                            sHol = global.hWEEKDAY;
                                        }

                                        // ②休日テーブルを参照し休日に該当するか調べます
                                        SysControl.SetDBConnect dc = new SysControl.SetDBConnect();
                                        OleDbCommand sCom2 = new OleDbCommand();
                                        sCom2.Connection = dc.cnOpen();
                                        OleDbDataReader dr = null;
                                        sCom2.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                                        sCom2.Parameters.Clear();
                                        sCom2.Parameters.AddWithValue("@year", global.sYear);
                                        sCom2.Parameters.AddWithValue("@Month", global.sMonth);
                                        sCom2.Parameters.AddWithValue("@day", sDays);
                                        dr = sCom2.ExecuteReader();
                                        if (dr.Read())
                                        {
                                            sHol = global.hHOLIDAY;
                                        }
                                        dr.Close();
                                        sCom2.Connection.Close();

                                        // ③休暇で判断
                                        // 振替休暇のとき
                                        if (stCSV[i + 5].Trim() == global.eFURIKYU)
                                        {
                                            sHol = global.hHOLIDAY; // 休日扱いとする

                                            //// 平日のときは振休とする（ＯＫ）
                                            //if (sHol != global.hSATURDAY && sHol != global.hHOLIDAY)
                                            //    sHol = global.hFURIKYU;
                                        }
                                        else if (stCSV[i + 5].Trim() == global.eFURIDE)　// 振替出勤のとき
                                        {
                                            sHol = global.hWEEKDAY; // 平日扱いとする

                                            //// 休日のときは振出とする（ＯＫ）
                                            //if (sHol == global.hSATURDAY || sHol == global.hHOLIDAY)
                                            //    sHol = global.hFURIDE;
                                            //else sHol = global.hWEEKDAY;    // 休日以外のときは平日扱いとする（ＮＧ）
                                        }

                                        sCom.Parameters.AddWithValue("@Hol", sHol);

                                        // テーブル書き込み
                                        sCom.ExecuteNonQuery();
                                    }
                                }
                                break;

                            case global.PART_SELECT:
                                sDays = 0;
                                for (int i = 14; i <= 314; i += 10)
                                {
                                    sCom.Parameters.Clear();
                                    sDays++;

                                    // 存在する日付のときにMDBへ登録する
                                    string tempDt = global.sYear + "/" + global.sMonth + "/" + sDays.ToString();
                                    if (DateTime.TryParse(tempDt, out dt))
                                    {
                                        // ヘッダID
                                        sCom.Parameters.AddWithValue("@ID", Utility.GetStringSubMax(stCSV[0], 17));

                                        // 日付
                                        sCom.Parameters.AddWithValue("@Days", sDays);

                                        // マーク
                                        sCom.Parameters.AddWithValue("@yukyu", Utility.GetStringSubMax(stCSV[i], 1));

                                        // 開始時間
                                        string hh = Utility.GetStringSubMax(stCSV[i + 1], 2);
                                        string mm = Utility.GetStringSubMax(stCSV[i + 2], 2);
                                        sCom.Parameters.AddWithValue("@sh", hh);
                                        sCom.Parameters.AddWithValue("@sm", mm);

                                        // 終了時間
                                        hh = Utility.GetStringSubMax(stCSV[i + 3], 2);
                                        mm = Utility.GetStringSubMax(stCSV[i + 4], 2);
                                        sCom.Parameters.AddWithValue("@eh", hh);
                                        sCom.Parameters.AddWithValue("@em", mm);

                                        // 休暇
                                        sCom.Parameters.AddWithValue("@kyukei", Utility.GetStringSubMax(stCSV[i + 5], 1));

                                        // 休憩なし
                                        sCom.Parameters.AddWithValue("@kyunashi", Utility.GetStringSubMax(stCSV[i + 6], 1));

                                        // 昼1
                                        sCom.Parameters.AddWithValue("@hiru1", stCSV[i + 7]);

                                        // 昼2
                                        sCom.Parameters.AddWithValue("@hiru2", stCSV[i + 8]);

                                        // 訂正
                                        sCom.Parameters.AddWithValue("@teisei", Utility.GetStringSubMax(stCSV[i + 9], 1));

                                        // 更新年月日
                                        sCom.Parameters.AddWithValue("@Date", DateTime.Today.ToShortDateString());

                                        // 休日区分の設定
                                        // ①曜日で判断します
                                        string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                                        int sHol = 0;
                                        if (Youbi == "土")
                                        {
                                            sHol = global.hSATURDAY;
                                        }
                                        else if (Youbi == "日")
                                        {
                                            sHol = global.hHOLIDAY;
                                        }
                                        else
                                        {
                                            sHol = global.hWEEKDAY;
                                        }

                                        // ②休日テーブルを参照し休日に該当するか調べます
                                        SysControl.SetDBConnect dc = new SysControl.SetDBConnect();
                                        OleDbCommand sCom2 = new OleDbCommand();
                                        sCom2.Connection = dc.cnOpen();
                                        OleDbDataReader dr = null;
                                        sCom2.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                                        sCom2.Parameters.Clear();
                                        sCom2.Parameters.AddWithValue("@year", global.sYear);
                                        sCom2.Parameters.AddWithValue("@Month", global.sMonth);
                                        sCom2.Parameters.AddWithValue("@day", sDays);
                                        dr = sCom2.ExecuteReader();
                                        if (dr.Read())
                                        {
                                            sHol = global.hHOLIDAY;
                                        }
                                        dr.Close();
                                        sCom2.Connection.Close();

                                        // ③休暇で判断
                                        // 振替休暇のとき
                                        if (stCSV[i + 5].Trim() == global.eFURIKYU)
                                        {
                                            sHol = global.hHOLIDAY;
                                        }
                                        else if (stCSV[i + 5].Trim() == global.eFURIDE)　// 振替出勤のときは平日扱い
                                        {
                                            sHol = global.hWEEKDAY;
                                        }

                                        sCom.Parameters.AddWithValue("@Hol", sHol);

                                        // テーブル書き込み
                                        sCom.ExecuteNonQuery();
                                    }
                                }
                                break;
                        }
                    }
                }

                // トランザクションコミット
                sTran.Commit();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                //CSVファイルを削除する
                foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.csv"))
                {
                    System.IO.File.Delete(files);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票CSVインポート処理", MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// MDBデータの件数をカウントする
        /// </summary>
        /// <returns></returns>
        private int CountMDB(int sKbn)
        {
            int rCnt = 0;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader dR;
            sCom.CommandText = "select ID from 勤務票ヘッダ where データ区分=? order by ID";
            sCom.Parameters.AddWithValue("@Kbn", sKbn);
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            return rCnt;
        }

        /// <summary>
        /// MDBデータのキー項目を配列に読み込む
        /// </summary>
        /// <returns></returns>
        private Entity.kData[] LoadMdbID(int sKbn)
        {
            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //レコード件数取得
            int cTotal = CountMDB(sKbn);

            Entity.kData[] DenID = new Entity.kData[1];

            int rCnt = 1;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader dR;
            sCom.CommandText = "select ID,所定勤務時間 from 勤務票ヘッダ where データ区分=? order by ID";
            sCom.Parameters.AddWithValue("@Kbn", sKbn);
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //プログレスバー表示
                frmP.Text = "勤務票データロード中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt / cTotal * 100;
                frmP.ProgressStep();

                //2件目以降は要素数を追加
                if (rCnt > 1)
                    DenID.CopyTo(DenID = new Entity.kData[rCnt], 0);

                DenID[rCnt - 1]._Meisai = new Entity.Gyou[_MULTIGYO];
                DenID[rCnt - 1]._sID = dR["ID"].ToString();
                DenID[rCnt - 1]._ShoTime = double.Parse(dR["所定勤務時間"].ToString());

                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            return DenID;
        }

        private void ErrInitial()
        {
            //エラー情報初期化
            lblErrMsg.Visible = false;
            global.errNumber = global.eNothing;     //エラー番号
            global.errMsg = string.Empty;           //エラーメッセージ
            lblErrMsg.Text = string.Empty;
        }

        private void DataShow(int sIx, Entity.kData[] rRec, DataGridView dgv)
        {
            // 出勤日数合計
            int rDays = 0;

            // 昼食回数
            int Lunchs = 0;

            string SqlStr = string.Empty;

            // 画像ファイル名
            global.pblImageFile = string.Empty;

            // データグリッドビュー初期化
            dataGridInitial(this.dg1);

            //データ表示背景色初期化
            dsColorInitial(this.dg1);

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            try
            {
                // 勤務票ヘッダ
                sCom.Connection = dCon.cnOpen();
                sCom.CommandText = "select * from 勤務票ヘッダ where ID = ?";
                sCom.Parameters.AddWithValue("@ID", rRec[sIx]._sID);

                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    txtID.Text = Utility.EmptytoZero(dR["シートID"].ToString());
                    txtYear.Text = Utility.EmptytoZero(dR["年"].ToString()).Substring(2, 2);
                    txtMonth.Text = Utility.EmptytoZero(dR["月"].ToString());
                    txtNo.Text = Utility.EmptytoZero(dR["個人番号"].ToString());

                    if (Utility.EmptytoZero(dR["勤務場所コード"].ToString()).Length > 8)
                    {
                        txtShozokuCode.Text = Utility.EmptytoZero(dR["勤務場所コード"].ToString()).Substring(1, 8);
                    }
                    else
                    {
                        txtShozokuCode.Text = Utility.EmptytoZero(dR["勤務場所コード"].ToString());
                    }

                    txtOrderCode.Text = Utility.EmptytoZero(dR["オーダーコード"].ToString());
                    txtID.Text = Utility.EmptytoZero(dR["シートID"].ToString());

                    if (Utility.NumericCheck(dR["離席時間2"].ToString()))
                    {
                        if (int.Parse(dR["離席時間2"].ToString()) > 0)
                        {
                            txtRiseki.Text = Utility.EmptytoZero(dR["離席時間1"].ToString()) + "." + Utility.EmptytoZero(dR["離席時間2"].ToString());
                        }
                        else
                        {
                            txtRiseki.Text = Utility.EmptytoZero(dR["離席時間1"].ToString());
                        }
                    }
                    else
                    {
                        txtRiseki.Text = Utility.EmptytoZero(dR["離席時間1"].ToString());
                    }

                    // パートタイマーもしくはスタッフで時間単位有休処理を行うとき：2017/03/16
                    if (_usrSel == global.PART_SELECT ||
                       (_usrSel == global.STAFF_SELECT && global.sTimeYukyu.ToString() == global.FLGON))
                    {
                        txtNichu.Text = Utility.EmptytoZero(dR["日中"].ToString());
                        txtChisou.Text = Utility.EmptytoZero(dR["遅早"].ToString());
                    }
                    else
                    {
                        txtNichu.Text = string.Empty;
                        txtChisou.Text = string.Empty;
                    }

                    txtShuTl.Text = Utility.EmptytoZero(dR["出勤日数合計"].ToString());
                    txtChuTl.Text = Utility.EmptytoZero(dR["昼食回数"].ToString());
                    if (dR["職場離脱"].ToString() == "0") chkRidatsu.Checked = false;
                    else chkRidatsu.Checked = true;
                    global.pblImageFile = dR["画像名"].ToString();

                    lblShoteiS.Text = dR["所定開始時間"].ToString();
                    lblShoteiE.Text = dR["所定終了時間"].ToString();
                    lblHanAmS.Text = dR["午前半休開始時間"].ToString();
                    lblHanAmE.Text = dR["午前半休終了時間"].ToString();
                    lblHanPmS.Text = dR["午後半休開始時間"].ToString();
                    lblHanPmE.Text = dR["午後半休終了時間"].ToString();

                    //////if (_usrSel == global.PART_SELECT)
                    //////{
                    //////    lblShoteiS.Text = dR["所定開始時間"].ToString();
                    //////    lblShoteiE.Text = dR["所定終了時間"].ToString();
                    //////    lblHanAmS.Text = dR["午前半休開始時間"].ToString();
                    //////    lblHanAmE.Text = dR["午前半休終了時間"].ToString();
                    //////    lblHanPmS.Text = dR["午後半休開始時間"].ToString();
                    //////    lblHanPmE.Text = dR["午後半休終了時間"].ToString();
                    //////}
                    //////else
                    //////{
                    //////    lblShoteiS.Text = string.Empty;
                    //////    lblShoteiE.Text = string.Empty;
                    //////    lblHanAmS.Text = string.Empty;
                    //////    lblHanAmE.Text = string.Empty;
                    //////    lblHanPmS.Text = string.Empty;
                    //////    lblHanPmE.Text = string.Empty;
                    //////}

                    //データ数表示
                    lblPage.Text = (_cI + 1).ToString() + "/" + inData.Length.ToString() + " 件目";
                }
                dR.Close();

                // マスターから個人情報を取得します
                if (_usrSel == global.STAFF_SELECT)
                    sCom.CommandText = "select * from スタッフマスタ where スタッフコード=?";
                else if (_usrSel == global.PART_SELECT)
                    sCom.CommandText = "select * from パートマスタ where 個人番号=?";

                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@sCode", txtNo.Text.Trim()).ToString();
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    switch (_usrSel)
                    {
                        case global.STAFF_SELECT:
                            lblName.Text = dR["スタッフ名"].ToString();
                            txtShozokuCode.Text = dR["派遣先CD"].ToString().Substring(1, 8);

                            if (int.Parse(dR["派遣先CD"].ToString().Substring(0, 1)) >= 2)
                                lblShozoku.Text = dR["派遣先名"].ToString() + " " + dR["派遣先部署"].ToString();
                            else lblShozoku.Text = dR["派遣先部署"].ToString();

                            break;

                        case global.PART_SELECT:
                            lblName.Text = dR["姓"].ToString() + " " + dR["名"].ToString();
                            txtShozokuCode.Text = dR["勤務場所店番"].ToString();
                            lblShozoku.Text = dR["勤務場所店名"].ToString();
                            break;
                    }
                }
                dR.Close();

                // 勤務記録明細
                StringBuilder sb = new StringBuilder();
                sb.Append("select 勤務票明細.*,勤務票ヘッダ.所定開始時間,勤務票ヘッダ.所定終了時間,勤務票ヘッダ.午前半休開始時間,勤務票ヘッダ.午前半休終了時間,勤務票ヘッダ.午後半休開始時間,勤務票ヘッダ.午後半休終了時間 ");
                sb.Append("from 勤務票明細 inner join 勤務票ヘッダ on 勤務票明細.ヘッダID = 勤務票ヘッダ.ID ");
                sb.Append("where ヘッダID = ? order by 勤務票明細.ID");
                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", rRec[sIx]._sID);
                dR = sCom.ExecuteReader();

                int r = 0;
                while (dR.Read())
                {
                    global.ShoS = dR["所定開始時間"].ToString();
                    global.ShoE = dR["所定終了時間"].ToString();
                    global.AmS = dR["午前半休開始時間"].ToString();
                    global.AmE = dR["午前半休終了時間"].ToString();
                    global.PmS = dR["午後半休開始時間"].ToString();
                    global.PmE = dR["午後半休終了時間"].ToString();

                    // ChangeValueイベント処理回避
                    global.dg1ChabgeValueStatus = false;

                    if (dR["マーク"].ToString() == global.FLGON) dgv[cMark, r].Value = global.MARU;
                    else dgv[cMark, r].Value = string.Empty;

                    // ChangeValueイベント処理実施
                    global.dg1ChabgeValueStatus = true;

                    if (dR["訂正"].ToString() == global.FLGON) dgv[cTeisei, r].Value = true;
                    else dgv[cTeisei, r].Value = false;

                    dgv[cKyuka, r].Value = dR["休暇"];
                    dgv[cSH, r].Value = dR["開始時"].ToString();
                    dgv[cSM, r].Value = dR["開始分"].ToString();
                    dgv[cEH, r].Value = dR["終了時"].ToString();
                    dgv[cEM, r].Value = dR["終了分"].ToString();

                    // ChangeValueイベント処理回避
                    global.dg1ChabgeValueStatus = false;

                    if (dR["休憩なし"].ToString() == global.FLGON) dgv[cKyukei, r].Value = global.MARU;
                    else dgv[cKyukei, r].Value = string.Empty;

                    if (_usrSel == global.PART_SELECT)
                    {
                        // 昼食
                        if (dR["昼1"].ToString() == global.FLGON) dgv[cHiru1, r].Value = global.MARU;
                        else dgv[cHiru1, r].Value = string.Empty;

                        if (dR["昼2"].ToString() == global.FLGON) dgv[cHiru2, r].Value = global.MARU;
                        else dgv[cHiru2, r].Value = string.Empty;

                        // 昼食回数加算
                        if (dR["昼1"].ToString() == global.FLGON || dR["昼2"].ToString() == global.FLGON)
                            Lunchs++;

                        // 出勤日数加算
                        if (dR["休日"].ToString() != global.hSATURDAY.ToString() &&
                            dR["休日"].ToString() != global.hHOLIDAY.ToString())
                        {
                            if (dR["マーク"].ToString() == global.FLGON ||
                                dR["開始時"].ToString() != string.Empty ||
                                dR["休暇"].ToString() == global.eAMHANKYU ||
                                dR["休暇"].ToString() == global.ePMHANKYU)
                                rDays++;
                        }
                    }
                    else
                    {
                        // 昼食
                        dgv[cHiru1, r].Value = string.Empty;
                        dgv[cHiru2, r].Value = string.Empty;

                        // 出勤日数加算
                        if (dR["マーク"].ToString() == global.FLGON ||
                            dR["開始時"].ToString() != string.Empty)
                            rDays++;
                    }

                    dgv[cID, r].Value = dR["ID"].ToString();

                    // ChangeValueイベント処理実施
                    global.dg1ChabgeValueStatus = true;

                    r++;

                    // データグリッド最下行まで達したら終了する（月末日以降のデータは無視する）
                    if (dgv.Rows.Count == r) break;
                }

                dR.Close();
                sCom.Connection.Close();

                // 離席時間テキストボックス表示色
                if (double.Parse(txtRiseki.Text) > 0) txtRiseki.BackColor = Color.LightPink;
                else txtRiseki.BackColor = Color.Empty;

                // 時間単位有休テキストボックス表示色
                if (Utility.StrToInt(txtNichu.Text) != 0) txtNichu.BackColor = Color.LightPink;
                else txtNichu.BackColor = Color.Empty;

                if (Utility.StrToInt(txtChisou.Text) != 0) txtChisou.BackColor = Color.LightPink;
                else txtChisou.BackColor = Color.Empty;

                //画像イメージ表示
                ShowImage(_InPath + global.pblImageFile);

                // ヘッダ情報
                txtYear.ReadOnly = false;
                txtMonth.ReadOnly = false;
                txtShozokuCode.ReadOnly = false;
                txtNo.ReadOnly = false;

                // スクロールバー設定
                hScrollBar1.Enabled = true;
                hScrollBar1.Minimum = 0;
                hScrollBar1.Maximum = rRec.Length - 1;
                hScrollBar1.Value = sIx;
                hScrollBar1.LargeChange = 1;
                hScrollBar1.SmallChange = 1;

                //最初のレコード
                if (sIx == 0)
                {
                    btnBefore.Enabled = false;
                    btnFirst.Enabled = false;
                }

                //最終レコード
                if ((sIx + 1) == rRec.Length)
                {
                    btnNext.Enabled = false;
                    btnEnd.Enabled = false;
                }

                //カレントセル選択状態としない
                dgv.CurrentCell = null;

                // その他のボタンを有効とする
                btnErrCheck.Enabled = true;
                btnDataMake.Enabled = true;
                btnDel.Enabled = true;

                // データグリッドビュー編集
                dg1.ReadOnly = false;

                //エラー情報表示
                ErrShow();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                global.dg1ChabgeValueStatus = true;
            }
        }

        //表示初期化
        private void dataGridInitial(DataGridView dgv)
        {
            txtID.Text = string.Empty;
            txtYear.Text = string.Empty;
            txtMonth.Text = string.Empty;
            txtNo.Text = string.Empty;
            txtShozokuCode.Text = string.Empty;
            txtOrderCode.Text = string.Empty;
            txtRiseki.Text = string.Empty;
            txtNichu.Text = string.Empty;
            txtChisou.Text = string.Empty;
            txtShuTl.Text = string.Empty;
            txtChuTl.Text = string.Empty;

            lblName.Text = string.Empty;
            lblShozoku.Text = string.Empty;
            lblShoteiS.Text = string.Empty;
            lblShoteiE.Text = string.Empty;
            lblHanAmS.Text = string.Empty;
            lblHanAmE.Text = string.Empty;
            lblHanPmS.Text = string.Empty;
            lblHanPmE.Text = string.Empty;

            txtID.BackColor = Color.Empty;
            txtYear.BackColor = Color.Empty;
            txtMonth.BackColor = Color.Empty;
            txtNo.BackColor = Color.Empty;
            txtShozokuCode.BackColor = Color.Empty;
            txtOrderCode.BackColor = Color.Empty;
            txtRiseki.BackColor = Color.Empty;
            txtNichu.BackColor = Color.Empty;
            txtChisou.BackColor = Color.Empty;
            txtShuTl.BackColor = Color.Empty;
            txtChuTl.BackColor = Color.Empty;

            txtID.ForeColor = Color.Navy;
            txtYear.ForeColor = Color.Navy;
            txtMonth.ForeColor = Color.Navy;
            txtNo.ForeColor = Color.Navy;
            txtShozokuCode.ForeColor = Color.Navy;
            txtOrderCode.ForeColor = Color.Navy;
            txtRiseki.ForeColor = Color.Navy;
            txtNichu.ForeColor = Color.Navy;
            txtChisou.ForeColor = Color.Navy;
            txtShuTl.ForeColor = Color.Navy;
            txtChuTl.ForeColor = Color.Navy;

            dgv.RowsDefaultCellStyle.ForeColor = Color.Navy;       //テキストカラーの設定
            dgv.DefaultCellStyle.SelectionBackColor = Color.Empty;
            dgv.DefaultCellStyle.SelectionForeColor = Color.Navy;

            //dgv.EditMode = EditMode.EditOnKeystrokeOrShortcutKey;

            //pictureBox1.Image = null;
            lblNoImage.Visible = false;
        }

        /// <summary>
        /// データ表示エリア背景色初期化
        /// </summary>
        /// <param name="dgv">データグリッドビューオブジェクト</param>
        private void dsColorInitial(DataGridView dgv)
        {
            // CellChangeValueイベント発生回避する
            global.dg1ChabgeValueStatus = false;

            txtID.BackColor = Color.White;
            txtYear.BackColor = Color.White;
            txtMonth.BackColor = Color.White;
            txtNo.BackColor = Color.White;
            txtShozokuCode.BackColor = Color.White;
            txtOrderCode.BackColor = Color.White;
            txtRiseki.BackColor = Color.White;
            txtNichu.BackColor = Color.White;
            txtChisou.BackColor = Color.White;
            txtShuTl.BackColor = Color.White;
            txtChuTl.BackColor = Color.Empty;

            // 行数
            dgv.RowCount = 0;

            // 行追加、日付セット
            //SysControl.SetDBConnect db = new SysControl.SetDBConnect();
            //OleDbCommand sCom = new OleDbCommand();
            //sCom.Connection = db.cnOpen();
            //OleDbDataReader dr = null;

            for (int i = 0; i < _MULTIGYO; i++)
            {
                DateTime dt;
                if (DateTime.TryParse(global.sYear.ToString() + "/" + global.sMonth.ToString() + "/" + (i + 1).ToString(), out dt))
                {
                    // 行を追加
                    dgv.Rows.Add();
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.Empty;

                    // 日
                    dgv[cDay, i].Value = i + 1;
                    // 曜日
                    string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                    dgv[cWeek, i].Value = Youbi;

                    //// 土日の場合
                    //if (Youbi == "日" || Youbi == "土")
                    //{
                    //    dgv.Rows[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    //}
                    //else
                    //{
                    //// 休日テーブルを参照し休日に該当するか調べます
                    //sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                    //sCom.Parameters.Clear();
                    //sCom.Parameters.AddWithValue("@year", global.sYear);
                    //sCom.Parameters.AddWithValue("@Month", global.sMonth);
                    //sCom.Parameters.AddWithValue("@day", i + 1);
                    //dr = sCom.ExecuteReader();
                    //if (dr.HasRows)
                    //{
                    //    dgv.Rows[i].DefaultCellStyle.BackColor = Color.MistyRose;
                    //}
                    //dr.Close();
                    //}

                    // 時分区切り記号
                    dgv[cS, i].Value = ":";
                    dgv[cE, i].Value = ":";
                }
            }
            //sCom.Connection.Close();

            // CellChangeValueイベントステータスもどす
            global.dg1ChabgeValueStatus = true;
        }

        /// <summary>
        /// 伝票画像表示
        /// </summary>
        /// <param name="iX">現在の伝票</param>
        /// <param name="tempImgName">画像名</param>
        public void ShowImage(string tempImgName)
        {
            string wrkFileName;

            //修正画面へ組み入れた画像フォームの表示    
            //画像の出力が無い場合は、画像表示をしない。
            if (tempImgName == string.Empty)
            {
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
                return;
            }

            //画像ファイルがあるときのみ表示
            wrkFileName = tempImgName;
            if (System.IO.File.Exists(wrkFileName))
            {
                leadImg.Visible = true;

                //画像ロード
                RasterCodecs.Startup();
                RasterCodecs cs = new RasterCodecs();

                // 描画時に使用される速度、品質、およびスタイルを制御します。 
                RasterPaintProperties prop = new RasterPaintProperties();
                prop = RasterPaintProperties.Default;
                prop.PaintDisplayMode = RasterPaintDisplayModeFlags.Resample;
                leadImg.PaintProperties = prop;

                leadImg.Image = cs.Load(wrkFileName, 0, CodecsLoadByteOrder.BgrOrGray, 1, 1);

                //画像表示倍率設定
                if (global.miMdlZoomRate == 0f)
                {
                    leadImg.ScaleFactor *= global.ZOOM_RATE;
                }
                else
                {
                    leadImg.ScaleFactor *= global.miMdlZoomRate;
                }

                //画像のマウスによる移動を可能とする
                leadImg.InteractiveMode = RasterViewerInteractiveMode.Pan;

                ////右へ90°回転させる
                //RotateCommand rc = new RotateCommand();
                //rc.Angle = 90 * 100;
                //rc.FillColor = new RasterColor(255, 255, 255);
                ////rc.Flags = RotateCommandFlags.Bicubic;
                //rc.Flags = RotateCommandFlags.Resize;
                //rc.Run(leadImg.Image);

                // グレースケールに変換
                GrayscaleCommand grayScaleCommand = new GrayscaleCommand();
                grayScaleCommand.BitsPerPixel = 8;
                grayScaleCommand.Run(leadImg.Image);
                leadImg.Refresh();

                cs.Dispose();
                RasterCodecs.Shutdown();
                global.pblImageFile = wrkFileName;

                lblNoImage.Visible = false;

                // 画像操作ボタン
                btnPlus.Enabled = true;
                btnMinus.Enabled = true;
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;
            }
            else
            {
                //画像ファイルがないとき
                leadImg.Visible = false;
                global.pblImageFile = string.Empty;
                lblNoImage.Visible = true;

                // 画像操作ボタン
                btnPlus.Enabled = false;
                btnMinus.Enabled = false;
                btnFirst.Enabled = true;
                btnNext.Enabled = true;
                btnBefore.Enabled = true;
                btnEnd.Enabled = true;
            }
        }

        private void btnFirst_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = 0;
            DataShow(_cI, inData, dg1);
            txtDNum.Text = string.Empty;
        }

        private void btnBefore_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (_cI > 0)
            {
                _cI--;
                DataShow(_cI, inData, dg1);
                txtDNum.Text = string.Empty;
            }
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            if (_cI + 1 < inData.Length)
            {
                _cI++;
                DataShow(_cI, inData, dg1);
                txtDNum.Text = string.Empty;
            }
        }

        private void btnEnd_Click(object sender, EventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = inData.Length - 1;
            DataShow(_cI, inData, dg1);
            txtDNum.Text = string.Empty;
        }

        private void txtRiseki_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '.')
            {
                e.Handled = true;
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
        ///  カレントデータの更新
        /// </summary>
        /// <param name="iX">カレントレコードのインデックス</param>
        private void CurDataUpDate(int iX)
        {
            //カレントデータを更新する
            string mySql = string.Empty;

            // エラーメッセージ
            string errMsg = string.Empty;

            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();

            //勤務票ヘッダテーブル
            mySql += "update 勤務票ヘッダ set ";
            mySql += "シートID=?,個人番号=?,年=?,月=?,勤務場所コード=?,オーダーコード=?,離席時間1=?,";
            mySql += "離席時間2=?,日中=?,遅早=?,職場離脱=?,出勤日数合計=?,昼食回数=?,更新年月日=?, ";
            mySql += "所定開始時間=?,所定終了時間=?,午前半休開始時間=?,午前半休終了時間=?,";
            mySql += "午後半休開始時間=?,午後半休終了時間=?,所定勤務時間=? ";
            mySql += "where ID = ?";

            errMsg = "勤務票ヘッダテーブル更新";

            sCom.CommandText = mySql;
            sCom.Parameters.AddWithValue("@SID", Utility.NulltoStr(txtID.Text));
            sCom.Parameters.AddWithValue("@No", Utility.NulltoStr(txtNo.Text).PadLeft(7, '0'));
            sCom.Parameters.AddWithValue("@Year", "20" + Utility.NulltoStr(txtYear.Text));

            //string mn = Utility.NulltoStr(txtMonth.Text);
            //if (Utility.NumericCheck(mn)) mn = int.Parse(mn).ToString();
            //sCom.Parameters.AddWithValue("@Month", mn);

            sCom.Parameters.AddWithValue("@Month", Utility.NulltoStr(txtMonth.Text));
            sCom.Parameters.AddWithValue("@Shozoku", Utility.NulltoStr(txtShozokuCode.Text));
            sCom.Parameters.AddWithValue("@Order", Utility.NulltoStr(txtOrderCode.Text));

            // 離席時間
            string[] sRi;
            if (txtRiseki.Text.Contains('.'))
            {
                // 離席時間ピリオド区切りで分割して配列に格納する
                sRi = txtRiseki.Text.Split('.');
                sCom.Parameters.AddWithValue("@Riseki1", sRi[0]);
                sCom.Parameters.AddWithValue("@Riseki2", sRi[1]);
            }
            else
            {
                sCom.Parameters.AddWithValue("@Riseki1", Utility.NulltoStr(txtRiseki.Text));
                sCom.Parameters.AddWithValue("@Riseki2", string.Empty);
            }

            sCom.Parameters.AddWithValue("@Nichu", Utility.NulltoStr(txtNichu.Text));
            sCom.Parameters.AddWithValue("@Chisou", Utility.NulltoStr(txtChisou.Text));

            if (chkRidatsu.Checked) sCom.Parameters.AddWithValue("@Ridatsu", global.FLGON);
            else sCom.Parameters.AddWithValue("@Ridatsu", global.FLGOFF);

            sCom.Parameters.AddWithValue("@ShuTl", Utility.NulltoStr(txtShuTl.Text));
            sCom.Parameters.AddWithValue("@ChuTl", Utility.NulltoStr(txtChuTl.Text));
            sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());
            sCom.Parameters.AddWithValue("@ShoS", lblShoteiS.Text);
            sCom.Parameters.AddWithValue("@ShoE", lblShoteiE.Text);
            sCom.Parameters.AddWithValue("@ShoAmS", lblHanAmS.Text);
            sCom.Parameters.AddWithValue("@ShoAmE", lblHanAmE.Text);
            sCom.Parameters.AddWithValue("@ShoPmS", lblHanPmS.Text);
            sCom.Parameters.AddWithValue("@ShoPmE", lblHanPmE.Text);
            sCom.Parameters.AddWithValue("@ShoTime", inData[_cI]._ShoTime);

            sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);

            sCom.Connection = dCon.cnOpen();

            // トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                sCom.ExecuteNonQuery();

                // 勤務票明細テーブル
                mySql = string.Empty;
                mySql += "update 勤務票明細 set ";
                mySql += "マーク=?,開始時=?,開始分=?,終了時=?,終了分=?,休暇=?,休憩なし=?,昼1=?,昼2=?,";
                mySql += "訂正=?,更新年月日=?,休日=? ";
                mySql += "where ID = ?";
                errMsg = "勤務票明細テーブル更新";
                sCom.CommandText = mySql;

                for (int i = 0; i < dg1.Rows.Count; i++)
                {
                    sCom.Parameters.Clear();

                    // マーク
                    if (dg1[cMark, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@mark", global.FLGON);
                    else sCom.Parameters.AddWithValue("@mark", global.FLGOFF);

                    // 開始時
                    sCom.Parameters.AddWithValue("@SH", Utility.NulltoStr(dg1[cSH, i].Value).ToString().Trim());
                    // 開始分
                    sCom.Parameters.AddWithValue("@SM", Utility.NulltoStr(dg1[cSM, i].Value).ToString().Trim());
                    // 終了時
                    sCom.Parameters.AddWithValue("@EH", Utility.NulltoStr(dg1[cEH, i].Value).ToString().Trim());
                    // 終了分
                    sCom.Parameters.AddWithValue("@EM", Utility.NulltoStr(dg1[cEM, i].Value).ToString().Trim());
                    // 休暇
                    sCom.Parameters.AddWithValue("@kyuka", Utility.NulltoStr(dg1[cKyuka, i].Value).ToString().Trim());

                    // 休憩なし
                    if (dg1[cKyukei, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@Kyukei", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Kyukei", global.FLGOFF);

                    // 昼１
                    if (dg1[cHiru1, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@Hiru1", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Hiru1", global.FLGOFF);

                    // 昼2
                    if (dg1[cHiru2, i].Value.ToString() == global.MARU)
                        sCom.Parameters.AddWithValue("@Hiru2", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Hiru2", global.FLGOFF);

                    // 訂正
                    if (dg1[cTeisei, i].Value.ToString() == "True")
                        sCom.Parameters.AddWithValue("@Teisei", global.FLGON);
                    else sCom.Parameters.AddWithValue("@Teisei", global.FLGOFF);

                    // 更新年月日
                    sCom.Parameters.AddWithValue("@date", DateTime.Today.ToShortDateString());

                    // 休日
                    sCom.Parameters.AddWithValue("@Hol", inData[iX]._Meisai[i]._clr);

                    // ID
                    sCom.Parameters.AddWithValue("@ID", Utility.NulltoStr(dg1[cID, i].Value).ToString());

                    // テーブル書き込み
                    sCom.ExecuteNonQuery();
                }

                //トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, errMsg, MessageBoxButtons.OK);

                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void btnPlus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor < global.ZOOM_MAX)
            {
                leadImg.ScaleFactor += global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void btnMinus_Click(object sender, EventArgs e)
        {
            if (leadImg.ScaleFactor > global.ZOOM_MIN)
            {
                leadImg.ScaleFactor -= global.ZOOM_STEP;
            }
            global.miMdlZoomRate = (float)leadImg.ScaleFactor;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!Utility.NumericCheck(txtDNum.Text))
            {
                MessageBox.Show("何件目のデータへ移動するか指定してください", "データ移動", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                return;
            }

            if (int.Parse(txtDNum.Text) > inData.Length)
            {
                MessageBox.Show("データ件数（" + inData.Length.ToString() + "件）の範囲を超えています", "データ移動", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                return;
            }

            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = int.Parse((txtDNum.Text).ToString()) - 1;
            DataShow(_cI, inData, dg1);
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            //カレントデータの更新
            CurDataUpDate(_cI);

            //エラー情報初期化
            ErrInitial();

            //レコードの移動
            _cI = hScrollBar1.Value;
            DataShow(_cI, inData, dg1);
            txtDNum.Text = string.Empty;
        }

        private void dg1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (e.Control is DataGridViewTextBoxEditingControl)
            {
                string ColName = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;

                if (ColName == cSH || ColName == cSM || ColName == cEH || ColName == cEM || ColName == cKyuka)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                    e.Control.KeyDown -= new KeyEventHandler(Control_KeyDown2);
                    //イベントハンドラを追加する
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                }

                if (ColName == cMark || ColName == cKyukei || ColName == cHiru1 || ColName == cHiru2)
                {
                    //イベントハンドラが複数回追加されてしまうので最初に削除する
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress);
                    e.Control.KeyPress -= new KeyPressEventHandler(Control_KeyPress2);
                    e.Control.KeyDown -= new KeyEventHandler(Control_KeyDown2);
                    //イベントハンドラを追加する
                    e.Control.KeyDown += new KeyEventHandler(Control_KeyDown2);
                    e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress2);
                }
            }
        }

        void Control_KeyDown2(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Space && e.KeyCode != Keys.Delete && e.KeyCode != Keys.Tab && e.KeyCode != Keys.Enter)
            {
                e.Handled = true;
                return;
            }
        }

        void Control_KeyPress2(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != (char)Keys.Space && e.KeyChar != '\b' && e.KeyChar != (char)Keys.Delete && e.KeyChar != (char)Keys.Tab && e.KeyChar != (char)Keys.Enter)
            {
                e.Handled = true;
                return;
            }
        }

        void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b' && e.KeyChar != '\t')
                e.Handled = true;
        }

        private void dg1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (global.dg1ChabgeValueStatus == false) return;
            global.dg1ChabgeValueStatus = false; // 自らChangeValueイベントを発生させない

            if (_cI == null) return;
            if (dg1.CurrentRow == null) return;

            // 該当カラム
            string Col = dg1.Columns[e.ColumnIndex].Name;

            // 該当行
            int Rn = e.RowIndex;

            if (Col == cDay || Col == cWeek || Col == cMark || Col == cS || Col == cE || Col == cHiru1 ||
                Col == cHiru2 || Col == cKyukei)
            {
                global.dg1ChabgeValueStatus = true;
                return;
            }

            global.lBackColorE = Color.FromArgb(251, 228, 183);
            global.lBackColorN = Color.FromArgb(255, 255, 255);

            // 休暇
            if (Col == cKyuka)
            {
                if (dg1[Col, Rn].Value != null)
                {
                    switch (dg1[Col, Rn].Value.ToString())
                    {
                        case "4":   // 振休
                            //// 営業日のとき 2014/01/27
                            //if (inData[_cI]._Meisai[Rn]._clr != 1 && inData[_cI]._Meisai[Rn]._clr != 2)
                            //{
                            //    inData[_cI]._Meisai[Rn]._clr = 2;
                            //}

                            inData[_cI]._Meisai[Rn]._clr = 2;   // 休日扱いとする
                            break;

                        case "3":   // 振出
                            //// 休日のとき 2014/01/27
                            //if (inData[_cI]._Meisai[Rn]._clr == 1 || inData[_cI]._Meisai[Rn]._clr == 2)
                            //{
                            //    inData[_cI]._Meisai[Rn]._clr = 0;
                            //    dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            //    dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                            //    dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            //}

                            inData[_cI]._Meisai[Rn]._clr = 0;   // 平日扱いとする
                            break;

                        default:
                            if (dg1[cWeek, Rn].Value.ToString() == "土")
                            {
                                inData[_cI]._Meisai[Rn]._clr = 1;
                            }
                            else if (dg1[cWeek, Rn].Value.ToString() == "日")
                            {
                                inData[_cI]._Meisai[Rn]._clr = 2;
                            }
                            else
                            {
                                inData[_cI]._Meisai[Rn]._clr = 0;
                            }

                            // 休日テーブルを参照し休日に該当するか調べます
                            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
                            OleDbCommand sCom = new OleDbCommand();
                            OleDbDataReader dr = null;
                            sCom.Connection = Con.cnOpen();
                            sCom.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                            sCom.Parameters.Clear();
                            sCom.Parameters.AddWithValue("@year", global.sYear);
                            sCom.Parameters.AddWithValue("@Month", global.sMonth);
                            sCom.Parameters.AddWithValue("@day", Rn + 1);
                            dr = sCom.ExecuteReader();
                            if (dr.Read())
                            {
                                inData[_cI]._Meisai[Rn]._clr = 2;
                            }
                            dr.Close();
                            sCom.Connection.Close();

                            break;
                    }
                }
            }

            switch (inData[_cI]._Meisai[Rn]._clr)
            {
                case 1: // 土曜日
                    global.lBackColorN = Color.FromArgb(225, 244, 255);
                    break;
                case 2: // 日曜日、カレンダー
                    global.lBackColorN = Color.FromArgb(255, 223, 227);
                    break;
                //case 4: // 振休
                //    global.lBackColorN = Color.FromArgb(255, 223, 227);
                //    break;
                default:
                    global.lBackColorN = Color.FromArgb(255, 255, 255);
                    break;
            }
            
            // 訂正チェック
            if (dg1[cTeisei, Rn].Value != null)
            {
                if (dg1[cTeisei, Rn].Value.ToString() == "True")
                {
                    global.lBackColorE = Color.FromArgb(251, 228, 183);
                    global.lBackColorN = Color.FromArgb(251, 228, 183);
                }

                // 行表示色
                TeiseiColor(Rn);
            }
            else
            {
                dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorN;
            }

            // 開始時間
            // 表示色をリセットします
            if (Col == cSH || Col == cSM)
            {
                dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
            }

            if ((Col == cSH || Col == cSM) && (dg1[cSH, Rn].Value != null && dg1[cSM, Rn].Value != null))
            {
                // 開始時
                if (Col == cSH)
                {
                    dg1[cSH, Rn].Style.BackColor = global.lBackColorN;

                    if (Utility.NumericCheck(dg1[cSH, Rn].Value.ToString()))
                        dg1[cSH, Rn].Value = dg1[cSH, Rn].Value.ToString().PadLeft(2, '0');

                    if (Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) < 8 && Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) != 0)
                        dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                    else dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                }

                // 開始分
                else if (Col == cSM)
                {
                    dg1[cSM, Rn].Style.BackColor = global.lBackColorN;

                    if (Utility.NumericCheck(dg1[cSM, Rn].Value.ToString()))
                    {
                        dg1[cSM, Rn].Value = dg1[cSM, Rn].Value.ToString().PadLeft(2, '0');

                        if (dg1[cSM, Rn].Value.ToString().Substring(1, 1) == "5" ||
                            dg1[cSM, Rn].Value.ToString().Substring(1, 1) == "0")
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                        else dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                    }
                }

                int hm = 0;

                // 所定開始時間と比較する
                if (dg1[cSH, Rn].Value.ToString() != string.Empty || dg1[cSM, Rn].Value.ToString() != string.Empty)
                {
                    // 所定開始時間を取得
                    hm = Utility.StrToInt(dg1[cSH, Rn].Value.ToString()) * 100 +
                        Utility.StrToInt(dg1[cSM, Rn].Value.ToString());

                    // 開始が所定より早い場合
                    if (Utility.StrToInt(global.ShoS.Replace(":", string.Empty)) > hm)
                    {
                        dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                    }

                    // 午後半休開始が所定より早い場合
                    if (dg1[cKyuka, Rn].Value.ToString() == "5" && (Utility.StrToInt(global.PmS.Replace(":", string.Empty)) > hm))
                    {
                        dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                    }

                    // 開始が所定より30分以上遅い場合
                    DateTime st;
                    if (DateTime.TryParse(global.ShoS, out st))
                    {
                        if (Utility.StrToInt(st.AddMinutes(30).ToShortTimeString().Replace(":", string.Empty)) <= hm)
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 休日出勤してる場合
                    if (inData[_cI]._Meisai[Rn]._clr == 1 || inData[_cI]._Meisai[Rn]._clr == 2 ||
                        inData[_cI]._Meisai[Rn]._clr == 3)
                    {
                        if (dg1[cSH, Rn].Value.ToString() != string.Empty || dg1[cSM, Rn].Value.ToString() != string.Empty)
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                        }
                        else
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                        }
                    }

                    // 半休で時間が掛かってる場合
                    if (dg1[cKyuka, Rn].Value.ToString() == "5")
                    {
                        if (Utility.StrToInt(global.AmS.Replace(":", string.Empty)) <= hm &&
                            Utility.StrToInt(global.AmE.Replace(":", string.Empty)) >= hm)
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    if (dg1[cKyuka, Rn].Value.ToString() == "6")
                    {
                        if (Utility.StrToInt(global.PmS.Replace(":", string.Empty)) <= hm &&
                            Utility.StrToInt(global.PmE.Replace(":", string.Empty)) >= hm)
                        {
                            dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }
                }
            }

            // 終了時間
            if (Col == cEH || Col == cEM)
            {
                dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
            }

            if ((Col == cEH || Col == cEM) && (dg1[cEH, Rn].Value != null && dg1[cEM, Rn].Value != null))
            {
                if (Col == cEH)
                {
                    if (Utility.NumericCheck(dg1[cEH, Rn].Value.ToString()))
                        dg1[cEH, Rn].Value = dg1[cEH, Rn].Value.ToString().PadLeft(2, '0');

                    if (Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) < 8 && Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) >= 22)
                        dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                    else dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                }

                // 終了分
                else if (Col == cEM)
                {
                    if (Utility.NumericCheck(dg1[cEM, Rn].Value.ToString()))
                    {
                        dg1[cEM, Rn].Value = dg1[cEM, Rn].Value.ToString().PadLeft(2, '0');

                        if (dg1[cEM, Rn].Value.ToString().Substring(1, 1) == "5" ||
                            dg1[cEM, Rn].Value.ToString().Substring(1, 1) == "0")
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                        else dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                    }
                }

                int hm = 0;

                if (dg1[cEH, Rn].Value.ToString() != string.Empty || dg1[cEM, Rn].Value.ToString() != string.Empty)
                {
                    // 終了時間を取得
                    hm = Utility.StrToInt(dg1[cEH, Rn].Value.ToString()) * 100 +
                        Utility.StrToInt(dg1[cEM, Rn].Value.ToString());

                    // 終了が所定より早い場合
                    if (Utility.StrToInt(global.ShoE.Replace(":", string.Empty)) > hm)
                    {
                        dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                    }

                    // 午後半休終了が所定より早い場合
                    if (dg1[cKyuka, Rn].Value.ToString() == "6" && (Utility.StrToInt(global.PmE.Replace(":", string.Empty)) > hm))
                    {
                        dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                        dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                    }

                    // 終了が所定より120分以上遅い場合
                    DateTime et;
                    if (DateTime.TryParse(global.ShoE, out et))
                    {
                        if (Utility.StrToInt(et.AddMinutes(120).ToShortTimeString().Replace(":", string.Empty)) <= hm)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 所定が15時以前でかつ終了が16時以降の場合
                    if (Utility.StrToInt(global.ShoE.Replace(":", string.Empty)) < 1600)
                    {
                        if (hm >= 1600)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    // 休日出勤してる場合
                    if (inData[_cI]._Meisai[Rn]._clr == 1 || inData[_cI]._Meisai[Rn]._clr == 2 ||
                        inData[_cI]._Meisai[Rn]._clr == 3)
                    {
                        if (dg1[cEH, Rn].Value.ToString() != string.Empty || dg1[cEM, Rn].Value.ToString() != string.Empty)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                        else
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                        }
                    }

                    // 半休で時間が掛かってる場合
                    if (dg1[cKyuka, Rn].Value.ToString() == "5")
                    {
                        if (Utility.StrToInt(global.AmS.Replace(":", string.Empty)) <= hm &&
                            Utility.StrToInt(global.AmE.Replace(":", string.Empty)) >= hm)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }

                    if (dg1[cKyuka, Rn].Value.ToString() == "6")
                    {
                        if (Utility.StrToInt(global.PmS.Replace(":", string.Empty)) <= hm &&
                            Utility.StrToInt(global.PmE.Replace(":", string.Empty)) >= hm)
                        {
                            dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                            dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                        }
                    }
                }
            }

            if (Col == cTeisei)
            {
                if (inData[_cI]._Meisai[Rn]._clr == 1)
                    global.lBackColorN = Color.FromArgb(225, 244, 255);
                else if (inData[_cI]._Meisai[Rn]._clr == 2)
                    global.lBackColorN = Color.FromArgb(255, 223, 227);
                else
                    global.lBackColorN = Color.FromArgb(255, 255, 255);
                
                // 行表示色
                TeiseiColor(Rn);
            }

            // ChangeValueイベントステータスを戻す
            global.dg1ChabgeValueStatus = true;
        }

        private void TeiseiColor(int Rn)
        {

            if (dg1[cTeisei, Rn].Value.ToString() == "True")
            {
                dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorE;
                dg1[cDay, Rn].Style.BackColor = global.lBackColorE;
                dg1[cWeek, Rn].Style.BackColor = global.lBackColorE;
                dg1[cMark, Rn].Style.BackColor = global.lBackColorE;
                dg1[cSH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cS, Rn].Style.BackColor = global.lBackColorE;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cEH, Rn].Style.BackColor = global.lBackColorE;
                dg1[cE, Rn].Style.BackColor = global.lBackColorE;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorE;
                dg1[cKyuka, Rn].Style.BackColor = global.lBackColorE;
                dg1[cKyukei, Rn].Style.BackColor = global.lBackColorE;
                dg1[cHiru1, Rn].Style.BackColor = global.lBackColorE;
                dg1[cHiru2, Rn].Style.BackColor = global.lBackColorE;
                dg1[cTeisei, Rn].Style.BackColor = global.lBackColorE;
            }
            else
            {
                dg1.Rows[Rn].DefaultCellStyle.BackColor = global.lBackColorN;
                dg1[cDay, Rn].Style.BackColor = global.lBackColorN;
                dg1[cWeek, Rn].Style.BackColor = global.lBackColorN;
                dg1[cMark, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cS, Rn].Style.BackColor = global.lBackColorN;
                dg1[cSM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEH, Rn].Style.BackColor = global.lBackColorN;
                dg1[cE, Rn].Style.BackColor = global.lBackColorN;
                dg1[cEM, Rn].Style.BackColor = global.lBackColorN;
                dg1[cKyuka, Rn].Style.BackColor = global.lBackColorN;
                dg1[cKyukei, Rn].Style.BackColor = global.lBackColorN;
                dg1[cHiru1, Rn].Style.BackColor = global.lBackColorN;
                dg1[cHiru2, Rn].Style.BackColor = global.lBackColorN;
                dg1[cTeisei, Rn].Style.BackColor = global.lBackColorN;
            }
        }

        private void dg1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            string colName = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;
            if (colName == cMark || colName == cKyukei || colName == cHiru1 || colName == cHiru2 || colName == cTeisei)
            {
                if (dg1.IsCurrentCellDirty)
                {
                    dg1.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    dg1.RefreshEdit();
                }
            }
        }

        private void dg1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string colName = dg1.Columns[e.ColumnIndex].Name;
            if (colName == cMark || colName == cHiru1 || colName == cHiru2 || colName == cKyukei)
            {
                if (dg1[colName, dg1.CurrentRow.Index].Value == null)
                {
                    dg1[colName, dg1.CurrentRow.Index].Value = global.MARU;
                }
                else if (dg1[colName, dg1.CurrentRow.Index].Value.ToString() == global.MARU)
                {
                    dg1[colName, dg1.CurrentRow.Index].Value = string.Empty;
                }
                else
                {
                    dg1[colName, dg1.CurrentRow.Index].Value = global.MARU;
                }

                dg1.RefreshEdit();

                // ChangeValueイベントステータスを戻す
                global.dg1ChabgeValueStatus = true; 
            }
        }

        private void dg1_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            // 入力値がSpaceのときに有効とします
            if (e.Value.ToString().Trim() != string.Empty) return;

            DataGridViewEx dgv = (DataGridViewEx)sender;
            string colName = dg1.Columns[e.ColumnIndex].Name;

            //セルの列を調べる
            
            if (colName == cMark || colName == cHiru1 || colName == cHiru2 || colName == cKyukei)
            {
                if (dgv[colName, dgv.CurrentRow.Index].Value == null ||
                    dgv[colName, dgv.CurrentRow.Index].Value.ToString() != global.MARU)
                {
                    //○をセルの値とする
                    e.Value = global.MARU;
                }
                else if (dgv[colName, dgv.CurrentRow.Index].Value.ToString() == global.MARU)
                {
                    //セルをEmptyとする
                    e.Value = string.Empty;
                }

                //解析が不要であることを知らせる
                e.ParsingApplied = true;

                // ChangeValueイベントステータスを戻す
                global.dg1ChabgeValueStatus = true; 
            }
        }

        private void dg1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            string ColH = string.Empty;
            string ColM = dg1.Columns[dg1.CurrentCell.ColumnIndex].Name;

            // 開始時間または終了時間を判断
            if (ColM == cSM)
            {
                ColH = cSH;
            }
            else if (ColM == cEM)
            {
                ColH = cEH;
            }
            else
            {
                return;
            }

            // 開始時、終了時が入力済みで開始分、終了分が未入力のとき"00"を表示します
            if (dg1[ColH, dg1.CurrentRow.Index].Value != null)
            {
                if (dg1[ColH, dg1.CurrentRow.Index].Value.ToString().Trim() != string.Empty)
                {
                    if (dg1[ColM, dg1.CurrentRow.Index].Value == null)
                    {
                        dg1[ColM, dg1.CurrentRow.Index].Value = "00";
                    }
                    else if (dg1[ColM, dg1.CurrentRow.Index].Value.ToString().Trim() == string.Empty)
                    {
                        dg1[ColM, dg1.CurrentRow.Index].Value = "00";
                    }
                }
            }
        }

        private void btnErrCheck_Click(object sender, EventArgs e)
        {
            //カレントレコード更新
            CurDataUpDate(_cI);

            //エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(inData[_cI]._sID, inData[inData.Length - 1]._sID) == false) return;

            //エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (_cI > 0)
            {
                if (ErrCheckMain(inData[0]._sID, inData[_cI - 1]._sID) == false) return;
            }

            // エラーなし
            DataShow(_cI, inData, dg1);
            MessageBox.Show("エラーはありませんでした", "エラーチェック", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dg1.CurrentCell = null;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     エラーチェックメイン処理 </summary>
        /// <param name="sID">
        ///     開始ID</param>
        /// <param name="eID">
        ///     終了ID</param>
        /// <returns>
        ///     True:エラーなし、false:エラーあり</returns>
        ///---------------------------------------------------------------
        private Boolean ErrCheckMain(string sIx, string eIx)
        {
            int rCnt = 0;

            //オーナーフォームを無効にする
            this.Enabled = false;

            //プログレスバーを表示する
            frmPrg frmP = new frmPrg();
            frmP.Owner = this;
            frmP.Show();

            //レコード件数取得
            int cTotal = CountMDB(_usrSel);

            //エラー情報初期化
            ErrInitial();

            //勤務記録表データ読み出し
            Boolean eCheck = true;
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;
            string mySql = string.Empty;

            mySql += "select * from 勤務票ヘッダ where データ区分=? order by ID";

            sCom.CommandText = mySql;
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@dkbn", _usrSel);
            sCom.Connection = dCon.cnOpen();
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;

                //プログレスバー表示
                frmP.Text = "エラーチェック実行中　" + rCnt.ToString() + "/" + cTotal.ToString();
                frmP.progressValue = rCnt * 100 / cTotal;
                frmP.ProgressStep();

                //指定範囲のIDならエラーチェックを実施する
                if (Int64.Parse(dR["ID"].ToString()) >= Int64.Parse(sIx) && Int64.Parse(dR["ID"].ToString()) <= Int64.Parse(eIx))
                {
                    eCheck = ErrCheckData(dR);
                    if (eCheck == false) break;　//エラーがあったとき
                }
            }

            dR.Close();
            sCom.Connection.Close();

            // いったんオーナーをアクティブにする
            this.Activate();

            // 進行状況ダイアログを閉じる
            frmP.Close();

            // オーナーのフォームを有効に戻す
            this.Enabled = true;

            //エラー有りの処理
            if (eCheck == false)
            {
                //エラーデータのインデックスを取得
                for (int i = 0; i < inData.Length; i++)
                {
                    if (inData[i]._sID == global.errID)
                    {
                        //エラーデータを画面表示
                        _cI = i;
                        DataShow(_cI, inData, dg1);
                        break;
                    }
                }
            }

            return eCheck;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     項目別エラーチェック </summary>
        /// <param name="cdR">
        ///     データリーダー</param>
        /// <returns>
        ///     エラーなし：true, エラー有り：false</returns>
        ///---------------------------------------------------------------
        private Boolean ErrCheckData(OleDbDataReader cdR)
        {
            // 昼食回数
            int Lunchs = 0;

            // 出勤日数
            int rDays = 0;

            // 休日フラグ 2014/01/16
            int HolVal = 0;

            //対象年
            if (cdR["年"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "年を入力してください";

                return false;
            }

            if (Utility.NumericCheck(cdR["年"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "数字を入力してください";

                return false;
            }

            if (global.sYear != int.Parse(cdR["年"].ToString()))
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "処理年と異なっています。 処理年月：" + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月";

                return false;
            }

            //対象月
            if (cdR["月"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eYear;
                global.errRow = 0;
                global.errMsg = "月を入力してください";

                return false;
            }

            if (Utility.NumericCheck(cdR["月"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "数字を入力してください";

                return false;
            }

            if (int.Parse(cdR["月"].ToString()) < 1 || int.Parse(cdR["月"].ToString()) > 12)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "正しい月を入力してください";

                return false;
            }

            if (global.sMonth != int.Parse(cdR["月"].ToString()))
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eMonth;
                global.errRow = 0;
                global.errMsg = "処理月と異なっています。 処理年月：" + global.sYear.ToString() + "年 " + global.sMonth.ToString() + "月";

                return false;
            }

            //個人番号
            //未入力のとき
            if (cdR["個人番号"] == null)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号を入力してください";

                return false;
            }

            if (cdR["個人番号"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号を入力してください";

                return false;
            }

            //数字以外のとき
            if (Utility.NumericCheck(cdR["個人番号"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号が不正です。";

                return false;
            }

            //個人番号マスター登録検査
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR;

            sCom.Connection = dCon.cnOpen();

            if (_usrSel == global.STAFF_SELECT)
                sCom.CommandText = "select * from スタッフマスタ where スタッフコード = ?";
            else if (_usrSel == global.PART_SELECT)
                sCom.CommandText = "select * from パートマスタ where 個人番号 = ?";

            sCom.Parameters.AddWithValue("@ID", cdR["個人番号"].ToString());
            dR = sCom.ExecuteReader();
            bool rHas = dR.HasRows;
            dR.Close();
            sCom.Connection.Close();

            if (!rHas)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eShainNo;
                global.errRow = 0;
                global.errMsg = "個人番号がマスターに存在しません";
                return false;
            }

            // シートID
            // 未入力のとき
            if (cdR["シートID"] == null)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "IDを入力してください";

                return false;
            }

            if (cdR["シートID"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "IDを入力してください";

                return false;
            }

            //数字以外のとき
            if (Utility.NumericCheck(cdR["シートID"].ToString()) == false)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "数字を入力してください";

                return false;
            }

            //範囲外のとき
            if (int.Parse(cdR["シートID"].ToString()) < 1 || int.Parse(cdR["シートID"].ToString()) > 4)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "IDが不正です";

                return false;
            }

            // 離席時間
            // 数字以外の入力のとき
            string riseki = cdR["離席時間1"].ToString() + cdR["離席時間2"].ToString();
            if (riseki != string.Empty)
            {
                if (Utility.NumericCheck(riseki) == false)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eRiseki;
                    global.errRow = 0;
                    global.errMsg = "離席時間が不正です";

                    return false;
                }
            }

            if (cdR["シートID"].ToString() == string.Empty)
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eID;
                global.errRow = 0;
                global.errMsg = "IDを入力してください";

                return false;
            }

            // 勤務記録明細データ
            sCom.Connection = dCon.cnOpen();
            sCom.CommandText = "select * from 勤務票明細 where ヘッダID = ? order by ID";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@HID", cdR["ID"].ToString());
            dR = sCom.ExecuteReader();

            //日付別データ
            int iX = 1;

            while (dR.Read())
            {
                // 休日か調べる 2014/01/27
                HolVal = 0;

                //////SysControl.SetDBConnect dCon2 = new SysControl.SetDBConnect();
                //////OleDbCommand sCom2 = new OleDbCommand();
                //////OleDbDataReader dr2;

                DateTime dt;
                if (DateTime.TryParse(global.sYear.ToString() + "/" + global.sMonth.ToString() + "/" + iX.ToString(), out dt))
                {
                    // 曜日
                    string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);

                    // 土日の場合
                    if (Youbi == "日" || Youbi == "土")
                    {
                        HolVal = 1;
                    }
                    else
                    {
                        // 休日テーブルを参照し休日に該当するか調べます
                        SysControl.SetDBConnect dCon2 = new SysControl.SetDBConnect();
                        OleDbCommand sCom2 = new OleDbCommand();
                        OleDbDataReader dr2;

                        sCom2.Connection = dCon2.cnOpen();
                        sCom2.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                        sCom2.Parameters.Clear();
                        sCom2.Parameters.AddWithValue("@year", global.sYear);
                        sCom2.Parameters.AddWithValue("@Month", global.sMonth);
                        sCom2.Parameters.AddWithValue("@day", iX);
                        dr2 = sCom2.ExecuteReader();
                        if (dr2.HasRows)
                        {
                            HolVal = 1;
                        }

                        dr2.Close();
                        sCom2.Connection.Close();
                    }
                }

                // エラーチェックの実行フラグ
                bool _errChek = true;

                // パートタイマーでエラーチェック対象外の行を判定します
                if (_usrSel == global.PART_SELECT)
                {
                    // 空白行のときエラーチェック対象外
                    if (dR["マーク"].ToString() == global.FLGOFF &&
                        dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty &&
                        dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty &&
                        dR["休暇"].ToString() == string.Empty && dR["休憩なし"].ToString() == global.FLGOFF &&
                        dR["昼1"].ToString() == global.FLGOFF && dR["昼2"].ToString() == global.FLGOFF && 
                        dR["訂正"].ToString() == global.FLGOFF)
                    {
                        _errChek = false;
                    }
                    ////// 休日以外で休暇欄(1,2,4,5,6)以外に記入がないときエラーチェック対象外 2014/01/09 「休日以外」の条件を追加
                    ////else if (dR["休日"].ToString() != global.hSATURDAY.ToString() &&
                    ////        dR["休日"].ToString() != global.hHOLIDAY.ToString() &&
                    ////        dR["休日"].ToString() != global.hFURIDE.ToString())

                    // 営業日で休暇欄(1,2,4,5,6)以外に記入がないときエラーチェック対象外 2014/01/09 「営業日」の条件を追加
                    else if (HolVal == 0)
                    {
                        if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU ||
                            dR["休暇"].ToString() == global.eFURIKYU ||
                            dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                        {
                            if (dR["マーク"].ToString() == global.FLGOFF &&
                                dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty &&
                                dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty &&
                                //dR["休憩なし"].ToString() == global.FLGOFF && 
                                dR["昼1"].ToString() == global.FLGOFF && dR["昼2"].ToString() == global.FLGOFF &&
                                dR["訂正"].ToString() == global.FLGOFF)
                            {
                                _errChek = false;
                            }
                        }
                    }
                }
                // スタッフでエラーチェック対象外の行を判定します
                else if (_usrSel == global.STAFF_SELECT)
                {
                    // 空白行のときエラーチェック対象外
                    if (dR["マーク"].ToString() == global.FLGOFF &&
                        dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty &&
                        dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty &&
                        dR["休暇"].ToString() == string.Empty && dR["休憩なし"].ToString() == global.FLGOFF &&
                        dR["訂正"].ToString() == global.FLGOFF)
                    {
                        _errChek = false;
                    }
                    //// 休日以外で休暇欄(1,2,4)以外に記入がないときエラーチェック対象外 2014/01/09 「休日以外」の条件を追加
                    //else if (dR["休日"].ToString() != global.hSATURDAY.ToString() &&
                    //        dR["休日"].ToString() != global.hHOLIDAY.ToString())

                    // 営業日で休暇欄(1,2,4)以外に記入がないときエラーチェック対象外 2014/01/09 「営業日」の条件を追加
                    else if (HolVal == 0)
                    {
                        if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU ||
                            dR["休暇"].ToString() == global.eFURIKYU)
                        {
                            if (dR["マーク"].ToString() == global.FLGOFF &&
                                dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty &&
                                dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty &&
                                //dR["休憩なし"].ToString() == global.FLGOFF && 
                                dR["昼1"].ToString() == global.FLGOFF && dR["昼2"].ToString() == global.FLGOFF &&
                                dR["訂正"].ToString() == global.FLGOFF)
                            {
                                _errChek = false;
                            }
                        }
                    }
                }

                // 明細行チェックの実施
                if (_errChek)
                {
                    if (!CheckMeisai(cdR, dR, iX, HolVal))
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errRow = iX - 1;
                        dR.Close();
                        sCom.Connection.Close();
                        return false;
                    }
                }

                // 昼食回数と出勤日数を加算します
                if (_usrSel == global.PART_SELECT)
                {
                    // 昼食回数加算
                    if (dR["昼1"].ToString() == global.FLGON || dR["昼2"].ToString() == global.FLGON)
                        Lunchs++;

                    // 出勤日数加算
                    if (dR["休日"].ToString() != global.hSATURDAY.ToString() &&
                        dR["休日"].ToString() != global.hHOLIDAY.ToString())
                    //if (HolVal == 0)
                    {
                        if (dR["マーク"].ToString() == global.FLGON ||
                            dR["開始時"].ToString() != string.Empty ||
                            dR["休暇"].ToString() == global.eAMHANKYU ||
                            dR["休暇"].ToString() == global.ePMHANKYU)
                            rDays++;
                    }
                }
                else
                {
                    // 出勤日数加算
                    if (dR["マーク"].ToString() == global.FLGON ||
                        dR["開始時"].ToString() != string.Empty)
                        rDays++;
                }

                iX++;
            }

            dR.Close();
            sCom.Connection.Close();

            // 出勤日数が記入回数と一致しているか
            if (cdR["出勤日数合計"].ToString().Trim() != rDays.ToString())
            {
                global.errID = cdR["ID"].ToString();
                global.errNumber = global.eDays;
                global.errRow = 0;
                global.errMsg = "出勤回数合計が記入された回数合計と不一致です。 計:" + rDays.ToString() + "日";
                return false;
            }

            // 昼食回数チェック
            if (_usrSel == global.PART_SELECT)
            {
                // 記入回数と一致しているか
                if (cdR["昼食回数"].ToString().Trim() != Lunchs.ToString())
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eLunch;
                    global.errRow = 0;
                    global.errMsg = "昼食回数合計が記入された回数合計と不一致です。計:" + Lunchs.ToString() + "回";
                    return false;
                }

                // 出勤日数内の値か
                if (Lunchs > rDays)
                {
                    global.errID = cdR["ID"].ToString();
                    global.errNumber = global.eLunch;
                    global.errRow = 0;
                    global.errMsg = "昼食回数合計が出勤回数合計を超えています";
                    return false;
                }
            }

            // スタッフ時間単位有休・日中、遅出・早帰
            if (global.sTimeYukyu == 1)
            {
                if (_usrSel == global.STAFF_SELECT)
                {
                    if (Utility.StrToInt(cdR["日中"].ToString()) > 40)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eNicchu;
                        global.errRow = 0;
                        global.errMsg = "日中時間が40時間を超えています";
                        return false;
                    }

                    if (Utility.StrToInt(cdR["遅早"].ToString()) > 40)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eOsode;
                        global.errRow = 0;
                        global.errMsg = "遅出・早帰時間が40時間を超えています";
                        return false;
                    }

                    if ((Utility.StrToInt(cdR["日中"].ToString()) +
                         Utility.StrToInt(cdR["遅早"].ToString())) > 40)
                    {
                        global.errID = cdR["ID"].ToString();
                        global.errNumber = global.eNicchu;
                        global.errRow = 0;
                        global.errMsg = "時間単位有休合計が40時間を超えています";
                        return false;
                    }
                }
            }

            return true;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     明細行毎のエラーチェック </summary>
        /// <param name="cdR">
        ///     勤務票ヘッダデータリーダー</param>
        /// <param name="dR">
        ///     勤務票明細データリーダー</param>
        /// <param name="hVal">
        ///     休日フラグ</param>
        /// <returns></returns>
        ///---------------------------------------------------------------
        private bool CheckMeisai(OleDbDataReader cdR, OleDbDataReader dR, int iX, int hVal)
        {
            TimeSpan ts;
            
            // 休日で休暇欄が空白か振出以外のときエラーとする 2014/01/09
            if (hVal == 1)
            {
                if (dR["休暇"].ToString() != string.Empty && dR["休暇"].ToString() != global.eFURIDE)
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休日の休暇欄が不正です";
                    return false;
                }
            }
            else
            {
                // 営業日で休暇欄が振出のときエラーとする  2014/01/27
                if (dR["休暇"].ToString() == global.eFURIDE)
                {
                    global.errNumber = global.eKyuka;
                    global.errMsg = "休暇欄が不正です";
                    return false;
                }
            }

            // パートタイマーのとき
            if (_usrSel == global.PART_SELECT)
            {
                // 休憩
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "半休で休憩なしが付いています";
                        return false;
                    }
                    else if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU || dR["休暇"].ToString() == global.eFURIKYU)
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "有給・特休・振休のいずれかで休憩なしが付いています";
                        return false;
                    }
                    else if (dR["昼1"].ToString() == "1")    // 2012/09/12 このチェックは旧バージョンになし
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "休憩なしで昼１が付いています";
                        return false;
                    }
                    else if (dR["昼2"].ToString() == "1")    // 2012/09/12 このチェックは旧バージョンになし
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "休憩なしで昼２が付いています";
                        return false;
                    }
                }

                // 昼１あり
                if (dR["昼1"].ToString() == "1")
                {
                    if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        global.errNumber = global.eHiru1;
                        global.errMsg = "半休で昼１が付いています";
                        return false;
                    }
                    else if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU || dR["休暇"].ToString() == global.eFURIKYU)
                    {
                        global.errNumber = global.eHiru1;
                        global.errMsg = "有給・特休・振休のいずれかで昼１が付いています";
                        return false;
                    }
                }

                // 昼２あり
                if (dR["昼2"].ToString() == "1")
                {
                    if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        global.errNumber = global.eHiru2;
                        global.errMsg = "半休で昼２が付いています";
                        return false;
                    }
                    else if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU || dR["休暇"].ToString() == global.eFURIKYU)
                    {
                        global.errNumber = global.eHiru2;
                        global.errMsg = "有給・特休・振休のいずれかで昼２が付いています";
                        return false;
                    }
                }

                if (dR["マーク"].ToString() == global.FLGON &&
                    dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty &&
                    dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty &&
                    dR["昼1"].ToString() == global.FLGOFF && dR["昼2"].ToString() == global.FLGOFF)
                {
                }
                else
                {
                    // 有給、休務のとき
                    if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU)
                    {
                        if (cdR["所定開始時間"].ToString() != (dR["開始時"].ToString().PadLeft(2, '0') + ":" + dR["開始分"].ToString().PadLeft(2, '0')))
                        {
                            global.errNumber = global.eSH;
                            global.errMsg = "有給で所定勤務時間が異なっています";
                            return false;
                        }
                        if (cdR["所定終了時間"].ToString() != (dR["終了時"].ToString().PadLeft(2, '0') + ":" + dR["終了分"].ToString().PadLeft(2, '0')))
                        {
                            global.errNumber = global.eEH;
                            global.errMsg = "有給で所定勤務時間が異なっています";
                            return false;
                        }
                    }
                }
            }
            // スタッフのとき
            else if (_usrSel == global.STAFF_SELECT)
            {
                // 休憩
                if (dR["休憩なし"].ToString() == "1")
                {
                    // 2017/03/13 半休のとき
                    if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "半休で休憩なしが付いています";
                        return false;
                    }
                    else if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU || dR["休暇"].ToString() == global.eFURIKYU)
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "有給・特休・振休のいずれかで休憩なしが付いています";
                        return false;
                    }
                }

                if (dR["マーク"].ToString() == global.FLGON &&
                    dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty &&
                    dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty &&
                    dR["昼1"].ToString() == global.FLGOFF && dR["昼2"].ToString() == global.FLGOFF)
                {
                }
                else
                {
                    // 有給、休務のとき
                    if (dR["休暇"].ToString() == global.eYUKYU || dR["休暇"].ToString() == global.eKYUMU)
                    {
                        if (cdR["所定開始時間"].ToString() != (dR["開始時"].ToString().PadLeft(2, '0') + ":" + dR["開始分"].ToString().PadLeft(2, '0')))
                        {
                            global.errNumber = global.eSH;
                            global.errMsg = "有給で所定勤務時間が異なっています";
                            return false;
                        }
                        if (cdR["所定終了時間"].ToString() != (dR["終了時"].ToString().PadLeft(2, '0') + ":" + dR["終了分"].ToString().PadLeft(2, '0')))
                        {
                            global.errNumber = global.eEH;
                            global.errMsg = "有給で所定勤務時間が異なっています";
                            return false;
                        }
                    }

                    // 半休のとき：2017/03/13
                    if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        // 開始・終了時刻
                        if (dR["開始時"].ToString() == string.Empty)
                        {
                            global.errNumber = global.eSH;
                            global.errMsg = "開始時刻を入力してください";
                            return false;
                        }

                        if (dR["開始分"].ToString() == string.Empty)
                        {
                            global.errNumber = global.eSM;
                            global.errMsg = "開始時刻を入力してください";
                            return false;
                        }

                        if (dR["終了時"].ToString() == string.Empty)
                        {
                            global.errNumber = global.eEH;
                            global.errMsg = "終了時刻を入力してください";
                            return false;
                        }

                        if (dR["終了分"].ToString() == string.Empty)
                        {
                            global.errNumber = global.eEM;
                            global.errMsg = "終了時刻を入力してください";
                            return false;
                        }

                        // マーク
                        if (dR["マーク"].ToString() == global.FLGON)
                        {
                            global.errNumber = global.eMark;
                            global.errMsg = "半休なので「○」は付けられません";
                            return false;
                        }

                        // 休日のとき
                        if (hVal == 1)
                        {
                            global.errNumber = global.eKyuka;
                            global.errMsg = "休日に半休は付けられません";
                            return false;
                        }
                    }
                }
            }


            // 振休だった場合
            if (dR["休暇"].ToString() == global.eFURIKYU)
            {
                if (dR["マーク"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "振替休日なので「○」は付けられません";
                    return false;
                }

                if (dR["開始時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["開始分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["終了時"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["終了分"].ToString() != string.Empty)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "振替休日なので時刻は入力できません";
                    return false;
                }

                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eKyukei;
                    global.errMsg = "振替休日なので休憩は付けられません";
                    return false;
                }

                if (dR["昼1"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru1;
                    global.errMsg = "振替休日なので昼１は付けられません";
                    return false;
                }

                if (dR["昼2"].ToString() == global.FLGON)
                {
                    global.errNumber = global.eHiru2;
                    global.errMsg = "振替休日なので昼2は付けられません";
                    return false;
                }
            }

            // 共通                   
            //if sKyuka = "9" And guDenData(giDenIndex).Head.sID = "3" Then
            //    If sMark != string.Empty || Trim(sSH) != string.Empty || Trim(sSM) != string.Empty || Trim(sEH) != string.Empty || Trim(sEM) != string.Empty || sHiru1 != string.Empty || sHiru2 != string.Empty Then
            //        sComment = "休暇欄が不正です。"
            //        guNgInfo.iKind = NG_KYUKA: bRet = False: GoTo NgProc
            //    Else
            //        GoTo NgProc
            //    End If
            //End If

            // 訂正欄チェック有り→エラーとする  2012/04/17
            if (dR["訂正"].ToString() == global.FLGON)
            {
                global.errNumber = global.eTeisei;
                global.errMsg = "訂正欄にチェックが入っています";
                return false;
            }

            if (dR["休暇"].ToString() == global.eYUKYU && dR["マーク"].ToString() == "1")
            {
                global.errNumber = global.eMark;
                global.errMsg = "有休なので○をはずしてください";
                return false;
            }

            // マークあり、休憩マークあり
            if (dR["マーク"].ToString() == "1" && dR["休憩なし"].ToString() == "1")
            {
                global.errNumber = global.eKyukei;
                global.errMsg = "休憩なしマークをはずしてください";
                return false;
            }

            // 開始時間 数字以外
            if (dR["開始時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["開始時"].ToString()))
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }

                // 開始時間 範囲外
                if (int.Parse(dR["開始時"].ToString()) < 0 || int.Parse(dR["開始時"].ToString()) > 23)
                {
                    global.errNumber = global.eSH;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }
            }

            // 開始分 数字以外
            if (dR["開始分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["開始分"].ToString()))
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }

                // 開始分 範囲外
                if (int.Parse(dR["開始分"].ToString()) < 0 || int.Parse(dR["開始分"].ToString()) > 59)
                {
                    global.errNumber = global.eSM;
                    global.errMsg = "開始時刻が不正です";
                    return false;
                }
            }
            
            // 終了時間 数字以外
            if (dR["終了時"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["終了時"].ToString()))
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }

                // 終了時間 範囲外
                if (int.Parse(dR["終了時"].ToString()) < 0 || int.Parse(dR["終了時"].ToString()) > 23)
                {
                    global.errNumber = global.eEH;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }
            }

            // 終了分 数字以外
            if (dR["終了分"].ToString() != string.Empty)
            {
                if (!Utility.NumericCheck(dR["終了分"].ToString()))
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }

                // 終了分 範囲外
                if (int.Parse(dR["終了分"].ToString()) < 0 || int.Parse(dR["終了分"].ToString()) > 59)
                {
                    global.errNumber = global.eEM;
                    global.errMsg = "終了時刻が不正です";
                    return false;
                }
            }

            // 振休以外　2014/01/27
            if (dR["休暇"].ToString() != global.eFURIKYU)
            {
                //  時間入力不備（マークがオンで開始時間、終了時間の何れかに記入が有る時） 2012/11/01
                if (dR["マーク"].ToString() == global.FLGON &&
                   (dR["開始時"].ToString() != string.Empty || dR["開始分"].ToString() != string.Empty ||
                    dR["終了時"].ToString() != string.Empty || dR["終了分"].ToString() != string.Empty))
                {
                    global.errNumber = global.eMark;
                    global.errMsg = "時間とマークが両方入力されています";
                    return false;
                }

                // 時間入力不備（①マークがオフ、②マークがオンで開始時間、終了時間の何れかに記入が有る時）
                if (dR["マーク"].ToString() == global.FLGOFF ||
                    (dR["マーク"].ToString() == global.FLGON &&
                    (dR["開始時"].ToString() != string.Empty || dR["開始分"].ToString() != string.Empty ||
                     dR["終了時"].ToString() != string.Empty || dR["終了分"].ToString() != string.Empty)))
                {
                    if (dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eSH;
                        global.errMsg = "開始時刻を入力してください";
                        return false;
                    }

                    if (dR["開始時"].ToString() == string.Empty && dR["開始分"].ToString() != string.Empty)
                    {
                        global.errNumber = global.eSH;
                        global.errMsg = "開始時刻を入力してください";
                        return false;
                    }

                    if (dR["開始時"].ToString() != string.Empty && dR["開始分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eSM;
                        global.errMsg = "開始時刻を入力してください";
                        return false;
                    }

                    if (dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "終了時刻を入力してください";
                        return false;
                    }

                    if (dR["終了時"].ToString() == string.Empty && dR["終了分"].ToString() != string.Empty)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "終了時刻を入力してください";
                        return false;
                    }

                    if (dR["終了時"].ToString() != string.Empty && dR["終了分"].ToString() == string.Empty)
                    {
                        global.errNumber = global.eEM;
                        global.errMsg = "終了時刻を入力してください";
                        return false;
                    }

                    // 開始時間と終了時間
                    int stm = int.Parse(dR["開始時"].ToString()) * 100 + int.Parse(dR["開始分"].ToString());
                    int etm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());

                    if (stm < 600)
                    {
                        global.errNumber = global.eSH;
                        global.errMsg = "開始時刻が6時以前になっています";
                        return false;
                    }

                    if (stm >= etm)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "終了時間が開始時間以前となっています";
                        return false;
                    }

                    // 勤務時間数
                    DateTime sHM = DateTime.Parse(dR["開始時"].ToString() + ":" + dR["開始分"].ToString());
                    DateTime eHM = DateTime.Parse(dR["終了時"].ToString() + ":" + dR["終了分"].ToString());

                    ts = eHM - sHM;
                    if ((ts.TotalMinutes <= 60) && (dR["休憩なし"].ToString() == global.FLGOFF))
                    {
                        global.errNumber = global.eSH;
                        global.errMsg = "勤務時間数がマイナスになっています";
                        return false;
                    }
                }
            }

            // 休暇
            if (dR["休暇"].ToString() != string.Empty)
            {
                if (_usrSel == global.PART_SELECT) // パートタイマー
                {
                    if (int.Parse(dR["休暇"].ToString()) < int.Parse(global.eYUKYU) ||
                        int.Parse(dR["休暇"].ToString()) > int.Parse(global.ePMHANKYU))
                    {
                        global.errNumber = global.eKyuka;
                        global.errMsg = "休暇欄が不正です";
                        return false;
                    }
                }
                else if (_usrSel == global.STAFF_SELECT) // スタッフ
                {
                    // 2017/03/13 半休を許可
                    if (int.Parse(dR["休暇"].ToString()) < int.Parse(global.eYUKYU) ||
                        int.Parse(dR["休暇"].ToString()) > int.Parse(global.ePMHANKYU))
                    {
                        global.errNumber = global.eKyuka;
                        global.errMsg = "休暇欄が不正です";
                        return false;
                    }
                }
            }

            // 昼食
            if (dR["昼1"].ToString() == global.FLGON && dR["昼2"].ToString() == global.FLGON)
            {
                global.errNumber = global.eHiru1;
                global.errMsg = "昼食が重複しています";
                return false;
            }

            // 休日だった場合
            if (dR["休暇"].ToString() == string.Empty)
            {
                if (hVal == 1)
                {
                    if (dR["マーク"].ToString() == global.FLGON)
                    {
                        global.errNumber = global.eMark;
                        global.errMsg = "休日出勤なので「○」は付けられません";
                        return false;
                    }

                    int tm = int.Parse(dR["終了時"].ToString()) * 100 + int.Parse(dR["終了分"].ToString());
                    if (tm > 2200)
                    {
                        global.errNumber = global.eEH;
                        global.errMsg = "休日出勤なので勤務終了は22時までです";
                        return false;
                    }

                    if (dR["昼1"].ToString() == global.FLGON)
                    {
                        global.errNumber = global.eHiru1;
                        global.errMsg = "休日出勤なので昼食は無効です";
                        return false;
                    }

                    if (dR["昼2"].ToString() == global.FLGON)
                    {
                        global.errNumber = global.eHiru2;
                        global.errMsg = "休日出勤なので昼食は無効です";
                        return false;
                    }
                }
            }

            // 勤務が6時間以上で休憩無しはエラー
            if (dR["開始時"].ToString() != string.Empty && dR["開始分"].ToString() != string.Empty &&
                dR["終了時"].ToString() != string.Empty && dR["終了分"].ToString() != string.Empty)
            {
                DateTime sh = DateTime.Parse(dR["開始時"].ToString() + ":" + dR["開始分"].ToString());
                DateTime eh = DateTime.Parse(dR["終了時"].ToString() + ":" + dR["終了分"].ToString());
                ts = Utility.GetTimeSpan(sh, eh);

                //if (ts.Hours >= 6)    2012/06/12 6時間を超過している場合に改修
                if (ts.Hours > 6 || (ts.Hours == 6 && ts.Minutes > 0))  // 2012/11/16 改修
                {
                    if (dR["休憩なし"].ToString() == global.FLGON)
                    {
                        global.errNumber = global.eKyukei;
                        global.errMsg = "勤務時間が6時間を超過して休憩が無しになっています";
                        return false;
                    }

                    if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                    {
                        global.errNumber = global.eKyuka;
                        global.errMsg = "半休で勤務時間が6時間を超過してます";
                        return false;
                    }
                }
            }

            // エラーなしを返す
            return true;
        }

        /// <summary>
        /// エラー情報表示
        /// </summary>
        private void ErrShow()
        {
            if (global.errNumber != global.eNothing)
            {
                lblErrMsg.Visible = true;
                lblErrMsg.Text = global.errMsg;

                //対象年
                if (global.errNumber == global.eYear)
                {
                    txtYear.BackColor = Color.Yellow;
                    txtYear.Focus();
                }

                //対象月
                if (global.errNumber == global.eMonth)
                {
                    txtMonth.BackColor = Color.Yellow;
                    txtMonth.Focus();
                }

                // シートID
                if (global.errNumber == global.eID)
                {
                    txtID.BackColor = Color.Yellow;
                    txtID.Focus();
                }

                //個人番号
                if (global.errNumber == global.eShainNo)
                {
                    txtNo.BackColor = Color.Yellow;
                    txtNo.Focus();
                }

                // 離席時間
                if (global.errNumber == global.eRiseki)
                {
                    txtRiseki.BackColor = Color.Yellow;
                    txtRiseki.Focus();
                }

                //マーク
                if (global.errNumber == global.eMark)
                {
                    dg1[cMark, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cMark, global.errRow];
                }

                // 開始時
                if (global.errNumber == global.eSH)
                {
                    dg1[cSH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cSH, global.errRow];
                }

                // 開始分
                if (global.errNumber == global.eSM)
                {
                    dg1[cSM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cSM, global.errRow];
                }

                // 終了時
                if (global.errNumber == global.eEH)
                {
                    dg1[cEH, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cEH, global.errRow];
                }

                // 終了分
                if (global.errNumber == global.eEM)
                {
                    dg1[cEM, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cEM, global.errRow];
                }

                // 休暇
                if (global.errNumber == global.eKyuka)
                {
                    dg1[cKyuka, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cKyuka, global.errRow];
                }

                // 休憩なし
                if (global.errNumber == global.eKyukei)
                {
                    dg1[cKyukei, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cKyukei, global.errRow];
                }

                // 昼１
                if (global.errNumber == global.eHiru1)
                {
                    dg1[cHiru1, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cHiru1, global.errRow];
                }

                // 昼２
                if (global.errNumber == global.eHiru2)
                {
                    dg1[cHiru2, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cHiru2, global.errRow];
                }
                
                // 訂正欄 2012/04/17
                if (global.errNumber == global.eTeisei)
                {
                    dg1[cTeisei, global.errRow].Style.BackColor = Color.Yellow;
                    dg1.Focus();
                    dg1.CurrentCell = dg1[cTeisei, global.errRow];
                }

                // 出勤日数
                if (global.errNumber == global.eDays)
                {
                    txtShuTl.BackColor = Color.Yellow;
                    txtShuTl.Focus();
                }

                // 昼食回数
                if (global.errNumber == global.eLunch)
                {
                    txtChuTl.BackColor = Color.Yellow;
                    txtChuTl.Focus();
                }
            }
        }

        private void dg1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnDataMake_Click(object sender, EventArgs e)
        {
            //カレントレコード更新
            CurDataUpDate(_cI);

            //エラーチェック実行①:カレントレコードから最終レコードまで
            if (ErrCheckMain(inData[_cI]._sID, inData[inData.Length - 1]._sID) == false) return;

            //エラーチェック実行②:最初のレコードからカレントレコードの前のレコードまで
            if (_cI > 0)
            {
                if (ErrCheckMain(inData[0]._sID, inData[_cI - 1]._sID) == false) return;
            }

            // 汎用データ作成
            if (MessageBox.Show("受け渡しデータを作成します。よろしいですか？", "勤怠データ登録", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            SaveData();
        }

        ///-----------------------------------------------------
        /// <summary>
        ///     受け渡しデータ作成 </summary>
        ///-----------------------------------------------------
        private void SaveData()
        {
            // 出力データ生成
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            OleDbDataReader dR = null;

            string OutFileName = string.Empty;
            try
            {
                //オーナーフォームを無効にする
                this.Enabled = false;

                //プログレスバーを表示する
                frmPrg frmP = new frmPrg();
                frmP.Owner = this;
                frmP.Show();

                //レコード件数取得
                int cTotal = CountMDBitem();
                int rCnt = 1;

                // データベース接続
                sCom.Connection = Con.cnOpen();
                StringBuilder sb = new StringBuilder();

                // ブレーク項目を定義
                string _DenID = string.Empty;   // ヘッダID
                string _kID = string.Empty;     // 個人番号

                // 月間集計処理を開始
                sb.Clear();

                if (_usrSel == global.PART_SELECT) // パート
                {
                    sb.Append("SELECT 勤務票ヘッダ.*,勤務票明細.*,パートマスタ.姓,パートマスタ.名 from ");
                    sb.Append("(勤務票ヘッダ inner join 勤務票明細 ");
                    sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID) ");
                    sb.Append("inner join パートマスタ ");
                    sb.Append("on 勤務票ヘッダ.個人番号 = パートマスタ.個人番号 ");
                    sb.Append("where データ区分=? ");
                    sb.Append("order by 勤務票ヘッダ.ID,勤務票明細.ID ");
                }
                else if (_usrSel == global.STAFF_SELECT) // スタッフ
                {
                    sb.Append("SELECT 勤務票ヘッダ.*,勤務票明細.*,スタッフマスタ.スタッフ名 from ");
                    sb.Append("(勤務票ヘッダ inner join 勤務票明細 ");
                    sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID) ");
                    sb.Append("inner join スタッフマスタ ");
                    sb.Append("on 勤務票ヘッダ.個人番号 = スタッフマスタ.スタッフコード ");
                    sb.Append("where データ区分=? ");
                    sb.Append("order by 勤務票ヘッダ.ID,勤務票明細.ID ");
                }

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@kbn", _usrSel);
                dR = sCom.ExecuteReader();

                // 出力ファイル名
                if (Properties.Settings.Default.PC == global.PC1SELECT)
                    OutFileName = "PC1_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒");
                else OutFileName = "PC2_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒");

                ////出力先フォルダがあるか？なければ作成する
                string outPath;
                if (_usrSel == global.PART_SELECT)
                {
                    OutFileName += ".txt";
                    outPath = global.sDAT2;
                }
                else
                {
                    OutFileName += ".csv";
                    outPath = global.sDAT;
                }
                
                if (!System.IO.Directory.Exists(outPath))
                    System.IO.Directory.CreateDirectory(outPath);

                // 出力ファイルインスタンス作成
                StreamWriter outFile = new StreamWriter(outPath + OutFileName, false, System.Text.Encoding.GetEncoding(932));
                StreamWriter outFile2 = null;

                // 集計値インスタンス生成
                Entity.saveData sd = new Entity.saveData();
                Entity.saveStaff ss = new Entity.saveStaff();

                if (_usrSel == global.PART_SELECT)
                {
                    SaveDataInitial(sd);
                }
                else if (_usrSel == global.STAFF_SELECT)
                {
                    SaveDataInitial(sd);
                    StaffDataInitial(ss);
                }

                bool sRFlg = false;
                int dIdx = 1;

                // Debug : 個人別出力ファイル
                StreamWriter debFile = null;

                // 明細書き出し
                while (dR.Read())
                {
                    //プログレスバー表示
                    frmP.Text = "汎用データ作成中です・・・" + rCnt.ToString() + "/" + cTotal.ToString();
                    frmP.progressValue = rCnt / cTotal * 100;
                    frmP.ProgressStep();

                    // スタッフで離脱があったら出力ストリームを生成する
                    if (_usrSel == global.STAFF_SELECT)
                    {
                        if (dR["職場離脱"].ToString() == global.FLGON && sRFlg == false)
                        {
                            // 出力ファイル名
                            if (Properties.Settings.Default.PC == global.PC1SELECT)
                                OutFileName = "PC1_R_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";
                            else OutFileName = "PC2_R_" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";

                            // 出力ファイルインスタンス作成
                            outFile2 = new StreamWriter(global.sDAT + OutFileName, false, System.Text.Encoding.GetEncoding(932));
                            sRFlg = true;
                        }
                    }

                    // ヘッダIDブレーク
                    if (_DenID != string.Empty && _DenID != dR["勤務票ヘッダ.ID"].ToString())
                    {
                        // CSV書き込み
                        if (_usrSel == global.PART_SELECT)
                        {
                            // 個人別CSVファイルに集計項目を書き込む
                            CsvWriteFooter(sd, debFile);

                            // 個人別出力CSVファイルを閉じる
                            debFile.Close();

                            // 受け渡しデータ作成
                            SaveDataToCsv(outFile, sd);
                        }
                        else if (_usrSel == global.STAFF_SELECT)
                        {
                            // 個人別CSVファイルに集計項目を書き込む
                            CsvWriteFooter(ss, debFile);

                            // 個人別出力CSVファイルを閉じる
                            debFile.Close();

                            // 受け渡しデータ作成
                            StaffDataToCsv(outFile, outFile2, ss, sd, sRFlg);
                        }

                        dIdx++;
                    }

                    if (_DenID != dR["勤務票ヘッダ.ID"].ToString())
                    {
                        string DebugPath = string.Empty;

                        // ヘッダIDを格納する
                        _DenID = dR["勤務票ヘッダ.ID"].ToString();

                        // 個人番号を格納する
                        _kID = dR["個人番号"].ToString();

                        // 集計クラス項目を初期化します
                        if (_usrSel == global.PART_SELECT)
                        {
                            // 出力ファイルインスタンス作成
                            DebugPath = Properties.Settings.Default.instPath + Properties.Settings.Default.OK +
                                dR["個人番号"].ToString() + " " + dR["姓"].ToString() + " " + dR["名"].ToString() + ".csv";

                            // 集計クラス項目の初期化
                            SaveDataInitial(sd);

                            // 時間単位有休休暇の値を取得します
                            sd.sNichu = Utility.StrToInt(dR["日中"].ToString());
                            sd.sChisou = Utility.StrToInt(dR["遅早"].ToString());

                            // 出勤日数の値を取得します
                            sd.sShuTl = Utility.StrToInt(dR["出勤日数合計"].ToString());
                        }
                        else if (_usrSel == global.STAFF_SELECT)
                        {
                            // 出力ファイルインスタンス作成
                            DebugPath = Properties.Settings.Default.instPath + Properties.Settings.Default.OK +
                                dR["個人番号"].ToString() + " " + dR["スタッフ名"].ToString() + ".csv";

                            // 集計クラス項目の初期化
                            SaveDataInitial(sd);
                            StaffDataInitial(ss);

                            // 時間単位有休処理 2017/03/15
                            if (global.sTimeYukyu.ToString() == global.FLGON)
                            {
                                // 時間単位有休休暇の値を取得します 2017/03/15
                                ss.sNichu = Utility.StrToInt(dR["日中"].ToString());
                                ss.sChisou = Utility.StrToInt(dR["遅早"].ToString());
                            }

                            // 出勤日数の値を取得します
                            ss.sShuTl = Utility.StrToInt(dR["出勤日数合計"].ToString());
                        }

                        // 個人別出力ファイルインスタンス作成
                        debFile = new StreamWriter(DebugPath, false, System.Text.Encoding.GetEncoding(932));

                        // 個人別ファイル表題行出力
                        debFile.WriteLine("日付,開始,終了,休暇,昼１,昼２,所定勤務時間,執務,残業");                            
                    }

                    // パートタイマー集計処理
                    if (_usrSel == global.PART_SELECT)
                    {
                        MakeSavePart(sd, dR, dIdx);

                        // 個人別CSVファイル出力
                        CsvWriteSub(dR, sd, debFile);
                    }
                    // スタッフ集計処理
                    else if (_usrSel == global.STAFF_SELECT)
                    {
                        MakeSaveStaff(ss, dR, dIdx);

                        // 個人別CSVファイル出力
                        CsvWriteSub(dR, ss, debFile);
                    }
                }

                // 受け渡しCSVを書き込み
                if (_usrSel == global.PART_SELECT)
                {
                    // 個人別CSVファイルに集計項目を書き込む
                    CsvWriteFooter(sd, debFile);

                    // 個人別出力CSVファイルを閉じる
                    debFile.Close();

                    // ＣＳＶ書き込み
                    SaveDataToCsv(outFile, sd);
                }
                else if (_usrSel == global.STAFF_SELECT)
                {
                    // 個人別CSVファイルに集計項目を書き込む
                    CsvWriteFooter(ss, debFile);

                    // 個人別出力CSVファイルを閉じる
                    debFile.Close();

                    // ＣＳＶ書き込み
                    StaffDataToCsv(outFile, outFile2, ss, sd, sRFlg);
                }

                // データリーダーをクローズします
                dR.Close();

                // 出力ファイルをクローズします
                outFile.Close();
                if (sRFlg == true) outFile2.Close();

                // いったんオーナーをアクティブにする
                this.Activate();

                // 進行状況ダイアログを閉じる
                frmP.Close();

                // オーナーのフォームを有効に戻す
                this.Enabled = true;

                // 画像ファイル退避先パスを取得します
                string tifPath =string.Empty;
                if (_usrSel == global.PART_SELECT) tifPath = global.sTIF2;
                else if (_usrSel == global.STAFF_SELECT) tifPath = global.sTIF;
                
                // 画像ファイルを退避します
                tifFileMove(tifPath);

                // 過去データ作成
                SaveLastData();

                // 勤務記録レコード削除
                DelItemRec(_usrSel);    // 勤務票明細データ
                DelHeadRec(_usrSel);    // 勤務票ヘッダデータ

                // スタッフ 設定月数分経過した過去画像を削除する
                delBackUpFiles(global.sBKDELS, global.sTIF);

                // パート 設定月数分経過した過去画像を削除する
                delBackUpFiles(global.sBKDELP, global.sTIF2);

                //終了
                MessageBox.Show("受け渡しデータが作成されました。", "勤怠データ作成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Tag = END_MAKEDATA;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (!dR.IsClosed) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();

                //MDBファイル最適化
                mdbCompact();
            }
        }

        /// <summary>
        /// 個人別のCSVファイルを出力
        /// </summary>
        /// <param name="dR">勤務票データリーダー</param>
        /// <param name="SD">集計値配列</param>
        /// <param name="deb">出力ストリーム</param>
        private void CsvWriteSub(OleDbDataReader dR, Entity.saveData SD, StreamWriter deb)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(dR["日付"].ToString().PadLeft(2, '0')).Append("日").Append(",");
                                
            // 開始・終了時間：マーク、有給、休務のときは所定勤務時間を適用します
            if ((dR["マーク"].ToString() == global.FLGON &&
                dR["開始時"].ToString() == string.Empty) ||
                (dR["休暇"].ToString() == global.eYUKYU ||
                dR["休暇"].ToString() == global.eKYUMU))
            {
                sb.Append(dR["所定開始時間"].ToString().PadLeft(5, '0')).Append(",");
                sb.Append(dR["所定終了時間"].ToString().PadLeft(5, '0')).Append(",");
            }
            else
            {
                sb.Append(dR["開始時"].ToString()).Append(":").Append(dR["開始分"].ToString()).Append(",");
                sb.Append(dR["終了時"].ToString()).Append(":").Append(dR["終了分"].ToString()).Append(",");
            }

            sb.Append(dR["休暇"].ToString()).Append(",");

            if (dR["昼1"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
            else sb.Append(string.Empty).Append(",");

            if (dR["昼2"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
            else sb.Append(string.Empty).Append(",");
                
            // 所定勤務時間
            sb.Append(SD.TimeJ.ToString()).Append(",");

            // 執務時間
            sb.Append(SD.sSitumu1.ToString()).Append(",");

            // 早出・残業
            sb.Append(SD.HayaZanSinya.ToString());

            // 出力します
            deb.WriteLine(sb.ToString());
        }

        private void CsvWriteSub(OleDbDataReader dR, Entity.saveStaff SD, StreamWriter deb)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(dR["日付"].ToString().PadLeft(2, '0')).Append("日").Append(",");
            
            // 勤務時間（所定時間）取得：マークのとき
            if (dR["マーク"].ToString() == global.FLGON && dR["開始時"].ToString() == string.Empty)
            {
                sb.Append(dR["所定開始時間"].ToString().PadLeft(5, '0')).Append(",");
                sb.Append(dR["所定終了時間"].ToString().PadLeft(5, '0')).Append(",");
            }

            // 勤務時間取得：時間記入のとき
            else if (dR["マーク"].ToString() == global.FLGOFF && dR["開始時"].ToString() != string.Empty)
            {
                sb.Append(dR["開始時"].ToString()).Append(":").Append(dR["開始分"].ToString()).Append(",");
                sb.Append(dR["終了時"].ToString()).Append(":").Append(dR["終了分"].ToString()).Append(",");
            }
            else
            {
                sb.Append(",");
                sb.Append(",");
            }

            sb.Append(dR["休暇"].ToString()).Append(",");

            if (dR["昼1"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
            else sb.Append(string.Empty).Append(",");

            if (dR["昼2"].ToString() == global.FLGON) sb.Append(global.MARU).Append(",");
            else sb.Append(string.Empty).Append(",");

            // 所定勤務時間
            sb.Append(SD.TimeJ.ToString()).Append(",");

            // 執務時間
            sb.Append(Utility.fncTime10(SD.sSitumu2).ToString()).Append(",");

            // 早出・残業
            sb.Append(SD.HayaZanSinya.ToString());

            // 出力します
            deb.WriteLine(sb.ToString());
        }

        /// <summary>
        /// 個人別CSVファイルに集計値を出力する
        /// </summary>
        /// <param name="SD">集計値配列</param>
        /// <param name="deb">出力ストリーム</param>
        private void CsvWriteFooter(Entity.saveData SD, StreamWriter deb)
        {
            deb.WriteLine("離席時間合計：" + SD.sRiseki.ToString() + "h");
            deb.WriteLine("時間単位有休　日中：" + SD.sNichu.ToString() + "h");
            deb.WriteLine("時間単位有休　遅出：" + SD.sChisou.ToString() + "h");

            if (SD.sRidatu.Trim() == global.FLGON)
                deb.WriteLine("職場離脱：" + global.MARU);
            else deb.WriteLine("職場離脱：");

            //deb.WriteLine("出勤日数：" + SD.sJituKinmu1.ToString() + "日");
            deb.WriteLine("出勤日数：" + SD.sShuTl.ToString() + "日");    // 2012/04/17 画面の出勤日数合計を出力する
            deb.WriteLine("昼食回数：" + (SD._sHiru1 + SD._sHiru2).ToString() + "回");
        }

        private void CsvWriteFooter(Entity.saveStaff SD, StreamWriter deb)
        {
            deb.WriteLine("離席時間合計：" + SD.sRiseki.ToString() + "h");

            // 時間単位有休処理を行うとき：2017/03/13
            if (global.sTimeYukyu.ToString() == global.FLGON)
            {
                deb.WriteLine("時間単位有休　日中：" + SD.sNichu.ToString() + "h");
                deb.WriteLine("時間単位有休　遅出：" + SD.sChisou.ToString() + "h");
            }
            else
            {
                deb.WriteLine("時間単位有休　日中：");
                deb.WriteLine("時間単位有休　遅出：");
            }

            if (SD.sRidatu.Trim() == global.FLGON)
                deb.WriteLine("職場離脱：" + global.MARU);
            else deb.WriteLine("職場離脱：");

            //deb.WriteLine("出勤日数：" + SD.dKinmuNisu.ToString() + "日");
            deb.WriteLine("出勤日数：" + SD.sShuTl.ToString() + "日");
            deb.WriteLine("昼食回数：" + (SD._sHiru1 + SD._sHiru2).ToString() + "回");
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     パートタイマー受け渡しデータ集計 </summary>
        /// <param name="SD">
        ///     集計データ</param>
        /// <param name="dR">
        ///     勤務票データリーダー</param>
        /// <param name="iX">
        ///     データインデックス</param>
        ///-------------------------------------------------------
        private void MakeSavePart(Entity.saveData SD, OleDbDataReader dR, int iX)
        {
            string _sSH = string.Empty;
            string _sSM = string.Empty;
            string _sEH = string.Empty;
            string _sEM = string.Empty;

            double TimeJ = 0;
            double TimeSHO = 0;

            DateTime TimeS = DateTime.Now;
            DateTime TimeE = DateTime.Now;
            double sShotei = 0;
            DateTime dt;
            DateTime sHayade = DateTime.Parse("8:30");

            SD.HayaZanSinya = 0;

            // 執務・残業・早出・深夜の計算領域を初期化します　2012/06/12
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.Zangyo = 0;
            SD.sSinya = 0;
            SD.Hayade = 0;
            SD.HayaZanSinya = 0; // 2012/07/31

            SD.sRecID = "200";
            SD.sCode = dR["個人番号"].ToString().PadLeft(7, '0');
            SD.sJinji = " ";
            SD.sNengetu = dR["年"].ToString() + dR["月"].ToString().PadLeft(2, '0');
            SD.sSeqNo = iX.ToString().PadLeft(5, '0');
            SD.sRidatu = dR["職場離脱"].ToString().PadLeft(2, ' ');

            // 離籍時間が記入されている場合
            if (dR["離席時間2"].ToString() != string.Empty)
                SD.sRiseki = double.Parse(Utility.StrToInt(dR["離席時間1"].ToString()).ToString() + "." + dR["離席時間2"].ToString());
            else SD.sRiseki = double.Parse(Utility.StrToInt(dR["離席時間1"].ToString()).ToString());

            // 所定開始終了時間を取得します
            // 午前半休のとき
            if (dR["休暇"].ToString() == global.eAMHANKYU)
            {
                if (DateTime.TryParse(dR["午後半休開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                if (DateTime.TryParse(dR["午後半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
            }
            // 午後半休のとき
            else if (dR["休暇"].ToString() == global.ePMHANKYU)
            {
                if (DateTime.TryParse(dR["午前半休開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                if (DateTime.TryParse(dR["午前半休終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
            }
            // その他
            else
            {
                if (DateTime.TryParse(dR["所定開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
                if (DateTime.TryParse(dR["所定終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
            }

            // 開始・終了時間を取得します
            _sSH = dR["開始時"].ToString();
            _sSM = dR["開始分"].ToString();
            _sEH = dR["終了時"].ToString();
            _sEM = dR["終了分"].ToString();

            // マーク、有給、休務のときは所定勤務時間を適用します
            if ((dR["マーク"].ToString() == global.FLGON &&
                dR["開始時"].ToString() == string.Empty) ||
                (dR["休暇"].ToString() == global.eYUKYU ||
                dR["休暇"].ToString() == global.eKYUMU))
            {
                _sSH = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sSM = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                _sEH = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sEM = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
            }

            // 実勤務日数を加算
            if (dR["マーク"].ToString() == global.FLGON || _sSH != string.Empty ||
                dR["休暇"].ToString() == global.eFURIDE || dR["休暇"].ToString() == global.eAMHANKYU ||
                dR["休暇"].ToString() == global.ePMHANKYU)
            {
                if ((dR["休日"].ToString() == global.hSATURDAY.ToString() ||
                     dR["休日"].ToString() == global.hHOLIDAY.ToString()) &&
                    (dR["休暇"].ToString() != global.eFURIDE))
                {
                }
                else SD.sJituKinmu1++;
            }

            // 実勤務日数 内有休日数
            if (dR["休暇"].ToString() == global.eYUKYU) SD.sJituKinmu2++;

            // 実勤務日数 内半休
            if (dR["休暇"].ToString() == global.eAMHANKYU ||
                dR["休暇"].ToString() == global.ePMHANKYU)
            {
                SD.sJituKinmu2 += 0.5;
                SD.sKinmuJikan2 += double.Parse(dR["所定勤務時間"].ToString());
            }

            // 実勤務日数 内特休日数
            if (dR["休暇"].ToString() == global.eKYUMU) SD.sJituKinmu3++;

            // 勤務時間 平常日所定勤務時間
            if (_sSH != string.Empty && _sSM != string.Empty && _sEH != string.Empty && _sEM != string.Empty)
            {
                // 開始時間を取得
                if (DateTime.TryParse(_sSH + ":" + _sSM, out dt)) TimeS = dt;

                // 終了時間を取得
                if (DateTime.TryParse(_sEH + ":" + _sEM, out dt)) TimeE = dt;

                // 勤務終了時間が早出開始時間以前の場合、勤務終了時間を早出開始時間とする
                if (TimeE < sHayade) sHayade = TimeE;

                // 開始時間のほうが大きいとき
                if (TimeS > TimeE)
                {
                    TimeSpan ts = new TimeSpan(24, 0, 0);
                    TimeE = TimeE + ts;
                }

                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                }
                else TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes - 60;

                // 休日だったら次の日へ
                if ((dR["休日"].ToString() == global.hSATURDAY.ToString() ||
                    dR["休日"].ToString() == global.hHOLIDAY.ToString()) &&
                    dR["休暇"].ToString() != global.eFURIDE)
                {
                    TimeJ = Utility.fncTime10(TimeJ);
                    SD.sKyujitu += TimeJ;

                    // 実勤務時間を取得
                    SD.TimeJ = TimeJ;

                    return;
                }

                // 
                // 所定時間内の勤務時間を取得します
                //
                // ①所定内勤務時間を取得します
                //if (SD.ShoteiE > TimeE)
                //    sShotei = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                //else sShotei = Utility.GetTimeSpan(TimeS, SD.ShoteiE).TotalMinutes;

                if (TimeE <= SD.ShoteiS || TimeS > SD.ShoteiE)
                    sShotei = 0;
                else
                {
                    // 開始時間
                    DateTime st;
                    if (TimeS < SD.ShoteiS) st = SD.ShoteiS;
                    else st = TimeS;

                    // 終了時間
                    DateTime et;
                    if (TimeE > SD.ShoteiE) et = SD.ShoteiE;
                    else et = TimeE;

                    // 所定時間内の勤務時間を取得します
                    sShotei = Utility.GetTimeSpan(st, et).TotalMinutes;
                }


                // ②休憩の有無、半休のとき
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    TimeJ = Utility.fncTime10(sShotei);
                    TimeSHO = sShotei;
                }
                // 半休のとき
                else if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    TimeJ = Utility.fncTime10(sShotei);
                    TimeSHO = sShotei;
                }
                else
                {
                    // 休憩有り60分減らす
                    TimeJ = Utility.fncTime10(sShotei - 60);
                    TimeSHO = sShotei - 60;
                }

                // ③平日のとき 実勤務時間に加算
                if (dR["休日"].ToString() != global.hSATURDAY.ToString() &&
                    dR["休日"].ToString() != global.hHOLIDAY.ToString())
                    SD.sKinmuJikan1 += TimeJ;
                else
                {
                    // ④休日に振替出勤のとき 実勤務時間に加算
                    if (dR["休暇"].ToString() == global.eFURIDE)
                        SD.sKinmuJikan1 += TimeJ;
                }

                // ⑤有休または特休取得日のとき「実勤務時間2」に加算
                if (dR["休暇"].ToString() == global.eYUKYU ||
                    dR["休暇"].ToString() == global.eKYUMU)
                    SD.sKinmuJikan2 += TimeJ;

                // ⑥半休のとき半休所定勤務時間（所定勤務時間の1/2）を実勤務時間に加算
                if (dR["休暇"].ToString() == global.eAMHANKYU ||
                    dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    TimeJ = double.Parse(dR["所定勤務時間"].ToString());
                    SD.sKinmuJikan1 += TimeJ;
                }
            }
            else
            {
                // 時間記入なしで半休のとき半休所定勤務時間（所定勤務時間の1/2）を実勤務時間に加算
                if (dR["休暇"].ToString() == global.eAMHANKYU ||
                    dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    TimeJ = double.Parse(dR["所定勤務時間"].ToString());
                    SD.sKinmuJikan1 += TimeJ;
                }
            }

            // 実勤務時間を取得
            SD.TimeJ = TimeJ;

            // 昼食回数１
            if (dR["昼1"].ToString() == global.FLGON) SD._sHiru1++;

            // 昼食回数２
            if (dR["昼2"].ToString() == global.FLGON) SD._sHiru2++;

            // "○"または開始時間が記入されているとき
            if (dR["マーク"].ToString() == global.FLGON || _sSH != string.Empty)
            {
                SD.TimeT = 0;
                SD.sSitumu2 = 0;

                //DateTime t830 = DateTime.Parse("8:30");

                //
                // 執務時間を設定します
                //
                // 開始時間が8:30以降のとき開始時間から9時間とします
                if (TimeS > sHayade)
                {
                    SD.SituS = TimeS;
                    SD.SituE = TimeS.AddMinutes(540);
                }
                else
                {
                    // 開始時間が8:30以前のとき8:30から9時間とします
                    SD.SituS = sHayade;
                    SD.SituE = sHayade.AddMinutes(540);
                }

                // 執務開始時間から執務終了時間または終了時間の早い時間までの経過時間を取得します
                TimeSpan ts;
                if (SD.SituE > TimeE)
                {
                    ts = Utility.GetTimeSpan(SD.SituS, TimeE);
                }
                else
                {
                    ts = Utility.GetTimeSpan(SD.SituS, SD.SituE);
                }

                if (ts.TotalMinutes > sShotei)  // 所定勤務時間以上のとき 
                {
                    SD.sSitumu1 = Utility.fncTime10(ts.TotalMinutes - sShotei);
                    SD.sSitumuT += SD.sSitumu1;
                }

                //    SD.Situ1 = ts.TotalMinutes - sShotei;
                //SD.sSitumu2 = SD.Situ1;
                //SD.sSitumuT += Utility.fncTime10(SD.sSitumu2);

                // ??????????????????????????????????????????????????????????
                if (DateTime.TryParse(dR["所定終了時間"].ToString(), out dt))
                {
                    SD.sETime = dt.AddMinutes(SD.sSitumu2);
                }
                // ??????????????????????????????????????????????????????????

                //
                // 平常日時間外 早出残業
                //
                // ①早出残業時間を取得します
                DateTime t840 = DateTime.Parse("8:40");
                if (TimeS < t840)
                {
                    if (TimeS <= sHayade)
                    {
                        SD.Hayade = Utility.GetTimeSpan(TimeS, sHayade).TotalMinutes;
                    }
                }

                // ②時間外残業
                if (TimeS < sHayade)
                {
                    if (dR["休憩なし"].ToString() == global.FLGON)
                        SD.ShoETime = sHayade.AddMinutes(480);
                    else SD.ShoETime = sHayade.AddMinutes(540);
                }
                else
                {
                    if (dR["休憩なし"].ToString() == global.FLGON)
                        SD.ShoETime = TimeS.AddMinutes(480);
                    else SD.ShoETime = TimeS.AddMinutes(540);
                }

                // 所定終了時間より終了時間が遅いとき 終了時間までの経過時間を取得します（但し22:00mまで）
                DateTime t2200 = DateTime.Parse("22:00");
                if (SD.ShoETime < TimeE)
                {
                    if (TimeE < t2200)
                        SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, TimeE).TotalMinutes;
                    else SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, t2200).TotalMinutes;
                }

                SD.HayaZanT += Utility.fncTime10(SD.Zangyo + SD.Hayade);

                // ③平常日時間外 深夜（22:00以降）
                if (TimeE > t2200)
                {
                    SD.sSinya = Utility.GetTimeSpan(t2200, TimeE).TotalMinutes;
                    SD.SinyaT += Utility.fncTime10(SD.sSinya);
                }

                SD.HayaZanSinya = Utility.fncTime10(SD.Zangyo + SD.Hayade + SD.sSinya);
            }
        }

        ///----------------------------------------------------
        /// <summary>
        ///     スタッフ受け渡しデータ集計 </summary>
        /// <param name="SD">
        ///     集計データ</param>
        /// <param name="dR">
        ///     勤務票データリーダー</param>
        /// <param name="iX">
        ///     データインデックス</param>
        ///----------------------------------------------------
        private void MakeSaveStaff(Entity.saveStaff SD, OleDbDataReader dR, int iX)
        {
            string _sSH = string.Empty;
            string _sSM = string.Empty;
            string _sEH = string.Empty;
            string _sEM = string.Empty;

            double TimeJ = 0;
            double TimeSHO = 0;

            DateTime TimeS = DateTime.Now;
            DateTime TimeE = DateTime.Now;
            double sShotei = 0;
            DateTime dt;
            bool Over360 = false;
            DateTime sHayade = DateTime.Parse("8:30");

            // 執務・残業計算領域を初期化します　2012/06/12
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.Zangyo = 0;
            SD.sSinya = 0;
            SD.Hayade = 0;
            SD.HayaZanSinya = 0; // 2012/07/31

            //// ID
            //switch (dR["シートID"].ToString())
            //{
            //    case "1":
            //        sHayade = DateTime.Parse("8:30");
            //        break;

            //    case "2":
            //        sHayade = DateTime.Parse("8:40");
            //        break;
            //    default:
            //        break;
            //}

            // ヘッダ項目
            SD.sCode =  dR["個人番号"].ToString();
            SD.sOrderCode = dR["オーダーコード"].ToString();
            SD.sKikanS = dR["年"].ToString() + dR["月"].ToString().PadLeft(2, '0') + "01";
            int lastDay = DateTime.DaysInMonth(int.Parse(dR["年"].ToString()), int.Parse(dR["月"].ToString()));
            SD.sKikanE = dR["年"].ToString() + dR["月"].ToString().PadLeft(2, '0') +  lastDay.ToString();
        
            SD.sRidatu = dR["職場離脱"].ToString();

            if (dR["離席時間2"].ToString() != string.Empty)
                SD.sRiseki = double.Parse(Utility.StrToInt(dR["離席時間1"].ToString()).ToString() + "." + dR["離席時間2"].ToString());
            else SD.sRiseki = double.Parse(Utility.StrToInt(dR["離席時間1"].ToString()).ToString());

            //// 所定開始終了時間を取得します
            //if (DateTime.TryParse(dR["所定開始時間"].ToString(), out dt)) SD.ShoteiS = dt;
            //if (DateTime.TryParse(dR["所定終了時間"].ToString(), out dt)) SD.ShoteiE = dt;
            
            // 所定開始終了時間を取得します：2017/03/13
            // 午前半休のとき
            if (dR["休暇"].ToString() == global.eAMHANKYU)
            {
                if (DateTime.TryParse(dR["午後半休開始時間"].ToString(), out dt))
                {
                    SD.ShoteiS = dt;
                }

                if (DateTime.TryParse(dR["午後半休終了時間"].ToString(), out dt))
                {
                    SD.ShoteiE = dt;
                }
            }
            // 午後半休のとき
            else if (dR["休暇"].ToString() == global.ePMHANKYU)
            {
                if (DateTime.TryParse(dR["午前半休開始時間"].ToString(), out dt))
                {
                    SD.ShoteiS = dt;
                }

                if (DateTime.TryParse(dR["午前半休終了時間"].ToString(), out dt))
                {
                    SD.ShoteiE = dt;
                }
            }
            // その他
            else
            {
                if (DateTime.TryParse(dR["所定開始時間"].ToString(), out dt))
                {
                    SD.ShoteiS = dt;
                }

                if (DateTime.TryParse(dR["所定終了時間"].ToString(), out dt))
                {
                    SD.ShoteiE = dt;
                }
            }
            
            // 勤務時間（所定時間）取得：マークのとき
            if (dR["マーク"].ToString() == global.FLGON && dR["開始時"].ToString() == string.Empty)
            {
                _sSH = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sSM = dR["所定開始時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
                _sEH = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(0, 2);
                _sEM = dR["所定終了時間"].ToString().Replace(":", string.Empty).Substring(2, 2);
            }

            // 勤務時間取得：時間記入のとき
            if (dR["マーク"].ToString() == global.FLGOFF && dR["開始時"].ToString() != string.Empty)
            {
                _sSH = dR["開始時"].ToString().PadLeft(2, '0');
                _sSM = dR["開始分"].ToString().PadLeft(2, '0');
                _sEH = dR["終了時"].ToString().PadLeft(2, '0');
                _sEM = dR["終了分"].ToString().PadLeft(2, '0');
            }

            // 勤務日数
            if (dR["マーク"].ToString() == global.FLGON || _sSH != string.Empty)
            {
                if ((dR["休日"].ToString() == global.hSATURDAY.ToString() || 
                     dR["休日"].ToString() == global.hHOLIDAY.ToString()) && 
                     dR["休暇"].ToString() != global.eFURIDE)
                {
                }
                else SD.dKinmuNisu++;
            }

            // 欠勤日数（ID=3のみ）
            //If guDenData(Denindex).Head.sID = "3" And .sKyuka = "9" Then
            //    KekkinNisu = KekkinNisu + 1
            //End If

            // 勤務時間
            if (_sSH != string.Empty && _sSM != string.Empty && _sEH != string.Empty && _sEM != string.Empty)
            {
                // 開始時間を取得
                if (DateTime.TryParse(_sSH + ":" + _sSM, out dt)) TimeS = dt;

                // 終了時間を取得
                if (DateTime.TryParse(_sEH + ":" + _sEM, out dt)) TimeE = dt;
                
                // 勤務終了時間が早出開始時間以前の場合、勤務終了時間を早出開始時間とする
                if (TimeE < sHayade) sHayade = TimeE;

                // 開始時間のほうが大きいとき
                if (TimeS > TimeE) 
                {
                    TimeSpan ts = new TimeSpan(24,0,0);
                    TimeE = TimeE + ts;
                }

                if (Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes > 360) Over360 = true;
                else Over360 = false;

                if (dR["休憩なし"].ToString() == global.FLGON)
                    TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                else TimeJ = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes - 60;

                // 休出日数
                if (dR["休日"].ToString() == global.hSATURDAY.ToString() || 
                    dR["休日"].ToString() == global.hHOLIDAY.ToString())
                {
                    if (dR["休暇"].ToString() != global.eFURIDE)
                    {
                        SD.sKyushutuNisu++;
                        TimeJ = Utility.fncTime10(TimeJ);
                        SD.sKyujitu += TimeJ;
                        SD.TimeJ = TimeJ;

                        // 平常日時間外 休日
                        SD.sShoteiKyujitu++;
                        return; // 休日だったら次の日へ
                    }
                }

                //// 休日だったら次の日へ
                //if (sColor = 1 Or .sColor = 2) And .sKyuka <> "3" Then
                //    TimeJ = fncTime10(CLng(TimeJ))
                //    sKyujitu = sKyujitu + TimeJ
                //    GoTo NextDayS
                //End If

                // 所定内の勤務時間を取得します
                //if (TimeE < SD.ShoteiE)
                //    sShotei = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                //else sShotei = Utility.GetTimeSpan(TimeS, SD.ShoteiE).TotalMinutes;

                
                if (TimeE <= SD.ShoteiS || TimeS > SD.ShoteiE) 
                    sShotei = 0;
                else 
                {
                    // 開始時間
                    DateTime st;
                    if (TimeS < SD.ShoteiS) st = SD.ShoteiS;
                    else st = TimeS;

                    // 終了時間
                    DateTime et;
                    if (TimeE > SD.ShoteiE) et = SD.ShoteiE;
                    else et = TimeE;

                    // 所定時間内の勤務時間を取得します
                    sShotei = Utility.GetTimeSpan(st, et).TotalMinutes;
                }

                //// 2017/0313
                //if (dR["休憩なし"].ToString() == global.FLGON)
                //{
                //    TimeJ = Utility.fncTime10(sShotei);
                //}
                //else
                //{
                //    TimeJ = Utility.fncTime10(sShotei - 60);
                //}
                
                // 休憩の有無、半休のとき：2017/03/13
                if (dR["休憩なし"].ToString() == global.FLGON)
                {
                    TimeJ = Utility.fncTime10(sShotei);
                }
                // 半休のとき
                else if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
                {
                    TimeJ = Utility.fncTime10(sShotei);
                }
                else
                {
                    // 休憩有り60分減らす
                    TimeJ = Utility.fncTime10(sShotei - 60);
                }
                
                if (dR["休日"].ToString() != global.hSATURDAY.ToString() &&
                    dR["休日"].ToString() != global.hHOLIDAY.ToString())
                {
                    SD.dKinmuJikan += TimeJ;
                }
                else
                {
                    if (dR["休暇"].ToString() == global.eFURIDE)
                    {
                        SD.dKinmuJikan += TimeJ;
                    }
                }

                // 6時間以内なら契約内へ加算
                if (!Over360)
                {
                    SD.Keiyakunai += TimeJ;
                }
            }

            // 勤務時間を取得します
            SD.TimeJ = TimeJ;

            // 有給日数・時間           
            if (dR["休暇"].ToString() == global.eYUKYU)
            {
                SD.YukyuNisu++;
                SD.YukyuJikan += Utility.fncTime10(Utility.GetTimeSpan(SD.ShoteiS,SD.ShoteiE).TotalMinutes - 60);
            }

            // 半休日数・時間：2017/03/13            
            if (dR["休暇"].ToString() == global.eAMHANKYU || dR["休暇"].ToString() == global.ePMHANKYU)
            {
                SD.YukyuNisu += 0.5;
                SD.YukyuJikan += Utility.fncTime10(Utility.GetTimeSpan(SD.ShoteiS, SD.ShoteiE).TotalMinutes);
            }

            // 特休日数・時間          
            if (dR["休暇"].ToString() == global.eKYUMU)
            {
                SD.TokukyuNisu++;
                SD.YukyuJikan +=  Utility.fncTime10(Utility.GetTimeSpan(SD.ShoteiS,SD.ShoteiE).TotalMinutes - 60);
            } 

            // 執務
            DateTime t2200 = DateTime.Parse("22:00");

            if (dR["マーク"].ToString() == global.FLGON || _sSH != string.Empty)
            {
                switch (dR["シートID"].ToString())
	            {                        
                    // 外部派遣者 ID=4の場合
                    case "4":
                        SD.TimeT = 0;
                        SD.sSitumu2 = 0;
                        SD.Situ1 = 0;
                        SD.Situ2 = 0;

                        ////// 所定時刻より早く出た場合
                        ////if (SD.ShoteiS > TimeS)
                        ////{
                        ////    // 終了時刻が設定早出時刻以降か 2012/04/24
                        ////    if (TimeE > SD.ShoteiS)
                        ////        SD.Situ1 = Utility.GetTimeSpan(TimeS, SD.ShoteiS).TotalMinutes;
                        ////    //else SD.Situ1 = Utility.GetTimeSpan(TimeS, TimeE).TotalMinutes;
                        ////}

                        //////if (SD.ShoteiS > TimeS)
                        //////    SD.Situ1 = Utility.GetTimeSpan(TimeS, SD.ShoteiS).TotalMinutes;
                        
                        ////if (dR["休憩なし"].ToString() == global.FLGON)
                        ////    SD.ShoETime = TimeS.AddMinutes(480);
                        ////else SD.ShoETime = TimeS.AddMinutes(540);
                                       
                        ////// 所定より遅く帰った場合で8時間以内
                        ////if (SD.ShoteiE < TimeE)
                        ////{
                        ////    if (SD.ShoETime > TimeE)
                        ////        SD.Situ2 = Utility.GetTimeSpan(SD.ShoteiE, TimeE).TotalMinutes;
                        ////    else SD.Situ2 = Utility.GetTimeSpan(SD.ShoteiE, SD.ShoETime).TotalMinutes;
                        ////}
                   
                        ////SD.sSitumu2 = SD.Situ2 + SD.Situ1;                    
                        ////if (SD.sSitumu2 < 0) SD.sSitumu2 = 0;
                        ////SD.sSitumuT += Utility.fncTime10(SD.sSitumu2);

                        //////sETime = Val(fncDateAdd(Val(ShoteiE), CInt(sSitumu2)))
                        

                        // 0424 ---------------------------------------------------------

                        //if (TimeS > SD.ShoteiS)
                        //{
                        //    SD.SituS = TimeS;
                        //    SD.SituE = TimeS.AddMinutes(540);
                        //}
                        //else
                        //{
                        //    SD.SituS = SD.ShoteiS;
                        //    SD.SituE = SD.ShoteiS.AddMinutes(540);
                        //}

                        SD.SituS = TimeS;
                        SD.SituE = TimeS.AddMinutes(540);

                        if (SD.SituE < TimeE)
                            SD.Situ2 = Utility.GetTimeSpan(SD.SituS, SD.SituE).TotalMinutes;
                        else SD.Situ2 = Utility.GetTimeSpan(SD.SituS, TimeE).TotalMinutes;
                    
                        if (SD.Situ2 > sShotei) SD.sSitumu2 = SD.Situ2 - sShotei;

                        SD.sSitumuT += Utility.fncTime10(SD.sSitumu2);

                        // ---------------------------------------------------------

                        if (dR["休憩なし"].ToString() == global.FLGON)
                            SD.ShoETime = TimeS.AddMinutes(480);
                        else SD.ShoETime = TimeS.AddMinutes(540);
                        
                        // 開始時間+(480 or 540)時間より終了時間が遅いとき残業時間を取得する
                        if (SD.ShoETime < TimeE)
                        {
                            if (TimeE < t2200)
                                SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, TimeE).TotalMinutes;
                            else SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, t2200).TotalMinutes;
                        }

                        SD.HayaZanT += Utility.fncTime10(SD.Zangyo + SD.Hayade);

                        // 平常日時間外 深夜
                        t2200 = DateTime.Parse("22:00");
                        if (TimeE > t2200)
                        {
                            SD.sSinya = Utility.GetTimeSpan(t2200, TimeE).TotalMinutes;
                            SD.SinyaT += Utility.fncTime10(SD.sSinya);
                        }

                        break;

                    // // ID=2、ID=3の場合
		            default:
                        // 平常日時間外 執務
                        SD.sSitumu2 = 0;
                        SD.Situ1 = 0;
                        SD.Situ2 = 0;

                        if (TimeS > sHayade)
                        {
                            SD.SituS = TimeS;
                            SD.SituE = TimeS.AddMinutes(540);
                        }
                        else
                        {
                            SD.SituS = sHayade;
                            SD.SituE = sHayade.AddMinutes(540);
                        }

                        if (SD.SituE < TimeE)
                            SD.Situ2 = Utility.GetTimeSpan(SD.SituS, SD.SituE).TotalMinutes;
                        else SD.Situ2 = Utility.GetTimeSpan(SD.SituS, TimeE).TotalMinutes;
                    
                        if (SD.Situ2 > sShotei) SD.sSitumu2 = SD.Situ2 - sShotei;

                        //If Situ2 - TimeSHO > 0 Then
                        //    Situ1 = Situ2 - TimeSHO
                        //End If
    
                        //sSitumu2 = Situ1 '+ Situ2

                        SD.sSitumuT += Utility.fncTime10(SD.sSitumu2);

                        //sETime = Val(fncDateAdd(Val(ShoteiE), CInt(sSitumu2)))

                        //TimeSHO = sShotei;

                        //r = 1
    
                        // 平常日時間外 早出残業
                        if (TimeS < sHayade)
                            SD.Hayade = Utility.GetTimeSpan(TimeS, sHayade).TotalMinutes;
    
                        if (TimeS < sHayade)
                        {
                            if (dR["休憩なし"].ToString() == global.FLGON)
                                SD.ShoETime = sHayade.AddMinutes(480);
                            else SD.ShoETime = sHayade.AddMinutes(540);
                        }
                        else
                        {
                            if (dR["休憩なし"].ToString() == global.FLGON)
                                SD.ShoETime = TimeS.AddMinutes(480);
                            else SD.ShoETime = TimeS.AddMinutes(540);
                        }

                        // 所定時間より終了時間が遅いとき残業時間を取得する
                        if (SD.ShoETime < TimeE)
                        {
                            t2200 = DateTime.Parse("22:00");
                            if (t2200 > TimeE)
                                SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, TimeE).TotalMinutes;
                            else SD.Zangyo = Utility.GetTimeSpan(SD.ShoETime, t2200).TotalMinutes;
                        }

                        //Do
                        //    iPastTime = Val(fncDateAdd(Val(ShoETime), r))
                        //    If Val(2200) >= iPastTime And iPastTime <= TimeE Then
                        //        Zangyo = Zangyo + 1
                        //        r = r + 1
                        //    Else
                        //        Exit Do
                        //    End If
                        //Loop
    
                        SD.HayaZanT += Utility.fncTime10(SD.Zangyo + SD.Hayade);

                        // 平常日時間外 深夜
                        t2200 = DateTime.Parse("22:00");
                        if (TimeE > t2200)
                        {
                            SD.sSinya = Utility.GetTimeSpan(t2200, TimeE).TotalMinutes;     
                            SD.SinyaT += Utility.fncTime10(SD.sSinya);    
                        }
                        break;
	            }
            }

            // 早出・残業・深夜を取得
            SD.HayaZanSinya = Utility.fncTime10(SD.Zangyo + SD.Hayade + SD.sSinya);

            //'ID=３の場合、所定日数を計算
            SD.dShoteiNisu++;

            //If .sYoubi <> "" Then
            //    SD. sShoteiNisu = sShoteiNisu + 1
            //End If

            // 平常日時間外 休日
            if (dR["休日"].ToString() == global.hSATURDAY.ToString() || 
                dR["休日"].ToString() == global.hHOLIDAY.ToString())
            {
                SD.sShoteiKyujitu++;
            }
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     パートタイマー受け渡しデータ集計項目初期化 </summary>
        /// <param name="SD">
        ///     受け渡しデータクラス</param>
        ///------------------------------------------------------------
        private void SaveDataInitial(Entity.saveData SD)
        {
            SD.sJikan1 = " 0";
            SD.sJikan2 = " 0";
            SD.sKinmu1 = "0";
            SD.sKinmu2 = "0";
            SD.sKinmu3 = "0";
            SD.sHiru1 = "0";
            SD.sHiru2 = "0";
            SD.sJigai1 = "0";
            SD.sJigai2 = "0";
            SD.sJigai3 = "0";
            SD.sJigai4 = "0";
            SD.P10 = "   0";
            SD.P13 = "   0";
            SD.P18 = "  0";
            SD.P19 = "  0";
            SD.P20 = "  0";
            SD.P21 = "  0";
            SD.P22 = " 0";
            SD.P23 = " 0";
            SD.P24 = " 0";
            SD.P25 = "  0";
            SD.P26 = "  0";
            SD.P27 = "  0";
            SD.P28 = "  0";
            SD.P29 = " 0";
            SD.P30 = " 0";
            SD.P31 = " 0";
            SD.P32 = " 0";
            SD.P34 = "  ";
            SD.uchi_yukyu_t = "    0";

            // 計算項目
            SD.sJituKinmu1 = 0;
            SD.sJituKinmu2 = 0;
            SD.sJituKinmu3 = 0;
            SD.sKinmuJikan1 = 0;
            SD.sKinmuJikan2 = 0;

            SD._sHiru1 = 0;
            SD._sHiru2 = 0;
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.sSitumuT = 0;
            SD.sHayaZan = 0;
            SD.sSinya = 0;
            SD.sKyujitu = 0;
            SD.Hayade = 0;
            SD.Zangyo = 0;
            SD.HayaZanT = 0;
            SD.HayaZanSinya = 0;
            SD.SinyaT = 0;
            SD.TimeT = 0;
            SD.Situ1 = 0;
            SD.Situ2 = 0;
            SD.sRiseki = 0;
            SD.sNichu = 0;
            SD.sChisou = 0;
            SD.TimeJ = 0;
            SD.sShuTl = 0;
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     スタッフ受け渡しデータ集計項目初期化 </summary>
        /// <param name="SD">
        ///     受け渡しデータクラス</param>
        ///------------------------------------------------------------
        private void StaffDataInitial(Entity.saveStaff SD)
        {
            // 計算項目
            SD.sJituKinmu1 = 0;
            SD.sJituKinmu2 = 0;
            SD.sJituKinmu3 = 0;
            SD.sKinmuJikan1 = 0;
            SD.sKinmuJikan2 = 0;

            SD._sHiru1 = 0;
            SD._sHiru2 = 0;
            SD.sSitumu1 = 0;
            SD.sSitumu2 = 0;
            SD.sSitumuT = 0;
            SD.sHayaZan = 0;
            SD.sSinya = 0;
            SD.sKyujitu = 0;
            SD.Hayade = 0;
            SD.Zangyo = 0;
            SD.HayaZanT = 0;
            SD.SinyaT = 0;
            SD.TimeT = 0;
            SD.Situ1 = 0;
            SD.Situ2 = 0;

            SD.dKinmuNisu = 0;
            SD.sKyushutuNisu = 0;
            SD.Keiyakunai = 0;
            SD.YukyuNisu = 0;
            SD.YukyuJikan = 0;
            SD.TokukyuNisu = 0;
            SD.sShoteiKyujitu = 0;
            SD.sRiseki = 0;
            SD.dKinmuJikan = 0;
            SD.dShoteiNisu = 0;
            SD.dKekkinNisu = 0;

            SD.sNichu = 0;
            SD.sChisou = 0;

            SD.TimeJ = 0;
            SD.HayaZanSinya = 0;
            SD.sShuTl = 0;

            SD.S77 = "";
            SD.S76 = "";
            SD.S83 = "";
            SD.S82 = "";
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     パートタイマー受け渡しデータを書き出す </summary>
        /// <param name="outFile">
        ///     出力先ストリーム</param>
        /// <param name="SD">
        ///     出力データクラス</param>
        ///------------------------------------------------------------
        private void SaveDataToCsv(StreamWriter outFile, Entity.saveData SD)
        {
            SD.sKinmu1 = SD.sJituKinmu1.ToString().PadLeft(2, ' ');

            // UCHI_YUKYU_D
            SD.sJituKinmu2 = SD.sJituKinmu2 * 1000;
            SD.sKinmu2 = (SD.sJituKinmu2 + ((SD.sNichu + SD.sChisou) * global.YUKYU0125)).ToString().PadLeft(4, '0').PadLeft(5, ' ');
            //SD.sKinmu2 = Utility.fncTen3(SD.sJituKinmu2 + ((SD.sNichu + SD.sChisou) * global.YUKYU0125)).PadLeft(4, '0').PadLeft(5, ' ');

            SD.sKinmu3 = SD.sJituKinmu3.ToString().PadLeft(2, ' ');

            // SYOTEI_H
            SD.sJikan1 = Utility.fncTen3(SD.sKinmuJikan1 - SD.sRiseki + SD.sChisou).PadLeft(4, ' ');

            // UCHI_YUKYU_H
            SD.sJikan2 = Utility.fncTen3(SD.sKinmuJikan2 + SD.sNichu + SD.sChisou).PadLeft(4, ' ');

            SD.sHiru1 = SD._sHiru1.ToString().PadLeft(2, ' ');
            SD.sHiru2 = SD._sHiru2.ToString().PadLeft(2, ' ');

            SD.sJigai1 = Utility.fncTen3(SD.sSitumuT).PadLeft(3, ' ');
            SD.sJigai2 = Utility.fncTen3(SD.HayaZanT).PadLeft(3, ' ');
            SD.sJigai3 = Utility.fncTen3(SD.SinyaT).PadLeft(3, ' ');
            SD.sJigai4 = Utility.fncTen3(SD.sKyujitu).PadLeft(3, ' ');

            double yt = (SD.sNichu + SD.sChisou) * global.YUKYU0125;
            if (yt > 0) SD.uchi_yukyu_t = yt.ToString().PadLeft(4, '0').PadLeft(5, ' ');
            else SD.uchi_yukyu_t = yt.ToString().PadLeft(5, ' ');

            // CSVファイルを書き出す
            StringBuilder sb = new StringBuilder();
            sb.Append(SD.sRecID);
            sb.Append(SD.sCode);
            sb.Append(SD.sJinji);
            sb.Append(SD.sNengetu);
            sb.Append(SD.sKinmu1);
            sb.Append(SD.sKinmu2);  // UCHI_YUKYU_D
            sb.Append(SD.sKinmu3);
            sb.Append(SD.sJikan1);  // SYOTEI_H
            sb.Append(SD.sJikan2);  // UCHI_YUKYU_H
            sb.Append(SD.P10);
            sb.Append(SD.sHiru1);
            sb.Append(SD.sHiru2);
            sb.Append(SD.P13);
            sb.Append(SD.sJigai1);
            sb.Append(SD.sJigai2);
            sb.Append(SD.sJigai3);
            sb.Append(SD.sJigai4);
            sb.Append(SD.P18);
            sb.Append(SD.P19);
            sb.Append(SD.P20);
            sb.Append(SD.P21);
            sb.Append(SD.P22);
            sb.Append(SD.P23);
            sb.Append(SD.P24);
            sb.Append(SD.P25);
            sb.Append(SD.P26);
            sb.Append(SD.P27);
            sb.Append(SD.P28);
            sb.Append(SD.P29);
            sb.Append(SD.P30);
            sb.Append(SD.P31);
            sb.Append(SD.P32);
            sb.Append(SD.sSeqNo);
            sb.Append(SD.sRidatu);
            sb.Append(SD.uchi_yukyu_t);     // UCHI_YUKYU_T 

            ////明細ファイル出力
            outFile.WriteLine(sb.ToString());
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     スタッフ受け渡しデータを書き出す </summary>
        /// <param name="outFile">
        ///     出力先ストリーム</param>
        /// <param name="SD">
        ///     出力データクラス</param>
        ///------------------------------------------------------------
        private void StaffDataToCsv(StreamWriter outFile, StreamWriter outFile2, Entity.saveStaff SD, Entity.saveData PT, bool sRFlg)
        {
            string OutFileName = string.Empty;
            StringBuilder sb = new StringBuilder();
            
            // 集計項目をセットします

            // 時間単位有休処理を行うとき　2017/03/15
            if (global.sTimeYukyu.ToString() == global.FLGON)
            {
                // 「日中」時間を「給与勤務時間」から当該時間を控除する。2017/03/15
                SD.sKinmuJikan = Utility.fncShosuZeroCut((double.Parse(Utility.fncTen2(SD.dKinmuJikan)) - SD.sRiseki - SD.sNichu).ToString("F1"));
               
                // 「日中」時間を「給与契約内」から当該時間を控除する。2017/03/15
                SD.sKeiyakunai = Utility.fncShosuZeroCut((double.Parse(Utility.fncTen2(SD.dKinmuJikan)) - SD.sRiseki - SD.sNichu).ToString("F1"));
                                
                if ((SD.sNichu + SD.sChisou) > 0)
                {
                    // 「日中」時間と「遅出早帰」時間の合計を「給与時単有休」にセット　2017/03/15
                    SD.S77 = Utility.fncShosuZeroCut((SD.sNichu + SD.sChisou).ToString("F1"));

                    // 同「給与時単回数」に「１」　2017/03/15
                    SD.S76 = "1";

                    // 「日中」と「遅出早帰」の合計を「請求時単有休」にセット　2017/03/15
                    SD.S83 = Utility.fncShosuZeroCut((SD.sNichu + SD.sChisou).ToString("F1"));
                    
                    // 同「請求時単回数」項目に「１」　2017/03/15
                    SD.S82 = "1";
                }
                else
                {
                    SD.S77 = "0";
                    SD.S76 = "0";
                    SD.S83 = "0";
                    SD.S82 = "0";
                }

                // 「日中」と「遅出早帰」の合計を「給与有休時間」「請求有休時間」項目に加算　2017/03/15
                SD.sYukyuJikan = Utility.fncTen2(SD.YukyuJikan + SD.sNichu + SD.sChisou);
            }
            else
            {
                SD.sKinmuJikan = Utility.fncShosuZeroCut((double.Parse(Utility.fncTen2(SD.dKinmuJikan)) - SD.sRiseki).ToString("F1"));
                SD.sKeiyakunai = Utility.fncShosuZeroCut((double.Parse(Utility.fncTen2(SD.dKinmuJikan)) - SD.sRiseki).ToString("F1"));
                SD.sYukyuJikan = Utility.fncTen2(SD.YukyuJikan);
            }

            SD.sKinmuNisu = SD.dKinmuNisu.ToString();
            SD.sYukyuNisu = SD.YukyuNisu.ToString();
            SD.sTokyu = SD.TokukyuNisu.ToString();
            SD.sKyushutu = SD.sKyushutuNisu.ToString();
            SD.sZangyo1 = Utility.fncTen2(SD.HayaZanT);
            SD.sZangyo2 = Utility.fncTen2(SD.sSitumuT);
            SD.sZangyo3 = Utility.fncTen2(SD.SinyaT);
            SD.sZangyo4 = Utility.fncTen2(SD.sKyujitu);
            SD.sShoteiNisu = (SD.dShoteiNisu - SD.sShoteiKyujitu).ToString();
            SD.sKekkin = SD.dKekkinNisu.ToString();
            
            // 離脱がありの場合は別データ出力
            if (SD.sRidatu == global.FLGON)
            {
                PT.sRecID = "   ";
                PT.sCode = SD.sCode.PadLeft(7,'0');
                PT.sJinji = " ";
                PT.sNengetu = global.sYear + global.sMonth.ToString().PadLeft(2, '0');
                PT.sRidatu = SD.sRidatu.ToString().PadLeft(2, ' ');
                PT.sKinmu1 = " 0";
                PT.sKinmu2 = "    0";
                PT.sKinmu3 = " 0";
                PT.sJikan1 = "   0";
                PT.sJikan2 = "   0";
                PT.sHiru1 = " 0";
                PT.sHiru2 = " 0";
    
                PT.sJigai1 = "  0";
                PT.sJigai2 = "  0";
                PT.sJigai3 = "  0";
                PT.sJigai4 = "  0";
                PT.sSeqNo = "    0";

                sb.Append(PT.sRecID);
                sb.Append(PT.sCode);
                sb.Append(PT.sJinji);
                sb.Append(PT.sNengetu);
                sb.Append(PT.sKinmu1);
                sb.Append(PT.sKinmu2);
                sb.Append(PT.sKinmu3);
                sb.Append(PT.sJikan1);
                sb.Append(PT.sJikan2);
                sb.Append(PT.P10);
                sb.Append(PT.sHiru1);
                sb.Append(PT.sHiru2);
                sb.Append(PT.P13);
                sb.Append(PT.sJigai1);
                sb.Append(PT.sJigai2);
                sb.Append(PT.sJigai3);
                sb.Append(PT.sJigai4);
                sb.Append(PT.P18);
                sb.Append(PT.P19);
                sb.Append(PT.P20);
                sb.Append(PT.P21);
                sb.Append(PT.P22);
                sb.Append(PT.P23);
                sb.Append(PT.P24);
                sb.Append(PT.P25);
                sb.Append(PT.P26);
                sb.Append(PT.P27);
                sb.Append(PT.P28);
                sb.Append(PT.P29);
                sb.Append(PT.P30);
                sb.Append(PT.P31);
                sb.Append(PT.P32);
                sb.Append(PT.sSeqNo);
                sb.Append(PT.sRidatu);
                sb.Append("    0");

                ////明細ファイル出力
                outFile2.WriteLine(sb.ToString());
            }

            // ID = 2 or ID = 4
            sb.Clear();
            sb.Append(SD.sCode).Append(",");
            sb.Append(SD.sOrderCode).Append(",");
            sb.Append(SD.sKikanS).Append(",");
            sb.Append(SD.sKikanE).Append(",");
            sb.Append(SD.sKinmuNisu).Append(",");
            sb.Append(SD.sKinmuJikan).Append(",");  // 給与勤務時間：日中時間を控除する
            sb.Append(SD.sYukyuNisu).Append(",");
            sb.Append(SD.sYukyuJikan).Append(",");　// 給与有休時間：「日中＋遅早」を加算
            sb.Append(SD.S09).Append(",");
            sb.Append(SD.S10).Append(",");
            sb.Append(SD.S11).Append(",");
            sb.Append(SD.sTokyu).Append(",");
            sb.Append(SD.sKyushutu).Append(",");
            sb.Append(SD.sKeiyakunai).Append(",");  // 給与計約内：日中時間を控除する
            sb.Append(SD.sZangyo1).Append(",");
            sb.Append(SD.sZangyo2).Append(",");
            sb.Append(SD.sZangyo3).Append(",");
            sb.Append(SD.sZangyo4).Append(",");
            sb.Append(SD.S19).Append(",");
            sb.Append(SD.S20).Append(",");
            sb.Append(SD.S21).Append(",");
            sb.Append(SD.S22).Append(",");
            sb.Append(SD.sKinmuNisu).Append(",");
            sb.Append(SD.sKinmuJikan).Append(",");
            sb.Append(SD.sYukyuNisu).Append(",");
            sb.Append(SD.sYukyuJikan).Append(",");  // 請求有休時間：「日中＋遅早」を加算
            sb.Append(SD.S27).Append(",");
            sb.Append(SD.S28).Append(",");
            sb.Append(SD.S29).Append(",");
            sb.Append(SD.sTokyu).Append(",");
            sb.Append(SD.sKyushutu).Append(",");
            sb.Append(SD.sKeiyakunai).Append(",");
            sb.Append(SD.sZangyo1).Append(",");
            sb.Append(SD.sZangyo2).Append(",");
            sb.Append(SD.sZangyo3).Append(",");
            sb.Append(SD.sZangyo4).Append(",");
            sb.Append(SD.S37).Append(",");
            sb.Append(SD.S38).Append(",");
            sb.Append(SD.S39).Append(",");
            sb.Append(SD.S40).Append(",");
            sb.Append(SD.S41).Append(",");
            sb.Append(SD.S42).Append(",");
            sb.Append(SD.S43).Append(",");
            sb.Append(SD.S44).Append(",");
            sb.Append(SD.S45).Append(",");
            sb.Append(SD.S46).Append(",");
            sb.Append(SD.S47).Append(",");
            sb.Append(SD.S48).Append(",");
            sb.Append(SD.S49).Append(",");
            sb.Append(SD.S50).Append(",");
            sb.Append(SD.S51).Append(",");
            sb.Append(SD.S52).Append(",");
            sb.Append(SD.S53).Append(",");
            sb.Append(SD.S54).Append(",");
            sb.Append(SD.S55).Append(",");
            sb.Append(SD.S56).Append(",");
            sb.Append(SD.S57).Append(",");
            sb.Append(SD.S58).Append(",");
            sb.Append(SD.S59).Append(",");
            sb.Append(SD.S60).Append(",");
            sb.Append(SD.S61).Append(",");
            sb.Append(SD.S62).Append(",");
            sb.Append(SD.S63).Append(",");
            sb.Append(SD.S64).Append(",");
            sb.Append(SD.S65).Append(",");
            sb.Append(SD.S66).Append(",");
            sb.Append(SD.S67).Append(",");
            sb.Append(SD.S68).Append(",");
            sb.Append(SD.S69).Append(",");
            sb.Append(SD.S70).Append(",");
            sb.Append(SD.S71).Append(",");
            sb.Append(SD.S72).Append(",");
            sb.Append(SD.S73).Append(",");
            sb.Append(SD.S74).Append(",");
            sb.Append(SD.S75).Append(",");
            sb.Append(SD.S76).Append(",");  // 給与時単回数：「日中＋遅早」>０のとき１加算
            sb.Append(SD.S77).Append(",");  // 給与時単有休：「日中＋遅早」を加算
            sb.Append(SD.S78).Append(",");
            sb.Append(SD.S79).Append(",");
            sb.Append(SD.S80).Append(",");
            sb.Append(SD.S81).Append(",");
            sb.Append(SD.S82).Append(",");  // 請求時単回数：「日中＋遅早」>０のとき１加算
            sb.Append(SD.S83);              // 請求時単有休：「日中＋遅早」を加算

            ////明細ファイル出力
            outFile.WriteLine(sb.ToString());
        }

        /// <summary>
        /// 過去データ作成
        /// </summary>
        private void SaveLastData()
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "select 個人番号 from 勤務票ヘッダ where データ区分 = ?";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@dkbn", _usrSel);
            
            OleDbCommand sCom2 = new OleDbCommand();
            sCom2.Connection = Con.cnOpen();
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("insert into 履歴 (");
            sb.Append("年,月,個人番号,氏名) values ");
            sb.Append("(?,?,?,?)");
            sCom2.CommandText = sb.ToString();

            OleDbDataReader dR = null;
            
            try 
            {
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@year", global.sYear.ToString());
                    sCom2.Parameters.AddWithValue("@month", global.sMonth.ToString());
                    sCom2.Parameters.AddWithValue("@Num", dR["個人番号"].ToString());
                    sCom2.Parameters.AddWithValue("@Name", string.Empty);
                    sCom2.ExecuteNonQuery();
	            }
	        }
	        catch (Exception ex)
	        {
                MessageBox.Show(ex.Message, "過去データ作成エラー", MessageBoxButtons.OK);
	        }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                if (sCom2.Connection.State == ConnectionState.Open) sCom2.Connection.Close();
            }
        }

        /// <summary>
        /// 勤務票明細レコード削除
        /// </summary>
        /// <param name="usr">データ区分（スタッフ：0,パートタイマー：1）</param>
        private void DelItemRec(int usr)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "select ID from 勤務票ヘッダ where データ区分 = ?";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@dkbn", usr);

            OleDbCommand sCom2 = new OleDbCommand();
            sCom2.Connection = Con.cnOpen();
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("delete from 勤務票明細 ");
            sb.Append("where ヘッダID=? ");
            sCom2.CommandText = sb.ToString();

            OleDbDataReader dR = null;

            try
            {
                dR = sCom.ExecuteReader();

                while (dR.Read())
                {
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@hID", dR["ID"].ToString());
                    sCom2.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票明細データ削除エラー", MessageBoxButtons.OK);
            }
            finally
            {
                if (dR.IsClosed == false) dR.Close();
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
                if (sCom2.Connection.State == ConnectionState.Open) sCom2.Connection.Close();
            }
        }

        /// <summary>
        /// 勤務票ヘッダレコード削除
        /// </summary>
        /// <param name="usr">データ区分（スタッフ：0,パートタイマー：1）</param>
        private void DelHeadRec(int usr)
        {
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();
            sCom.CommandText = "delete from 勤務票ヘッダ where データ区分 = ?";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@dkbn", usr);

            try
            {
                sCom.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票ヘッダデータ削除エラー", MessageBoxButtons.OK);
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        /// <summary>
        /// 画像ファイル退避処理
        /// </summary>
        private void tifFileMove(string tifPath)
        {
            //移動先フォルダがあるか？なければ作成する（TIFフォルダ）
            if (!System.IO.Directory.Exists(tifPath)) System.IO.Directory.CreateDirectory(tifPath);

            //画像を退避先フォルダへ移動する            
            foreach (string files in System.IO.Directory.GetFiles(_InPath, "*.tif"))
            {
                File.Move(files, tifPath + @"\" + System.IO.Path.GetFileName(files));
            }
        }

        /// <summary>
        /// 設定月数分経過した過去画像を削除する    
        /// </summary>
        private void imageDelete(int sdel)
        {
            //削除月設定が0のとき、「過去画像削除しない」とみなし終了する
            if (sdel == 0) return;

            try
            {
                //削除年月の取得
                DateTime delDate = DateTime.Today.AddMonths(sdel * (-1));
                int _dYY = delDate.Year;            //基準年
                int _dMM = delDate.Month;           //基準月
                int _dYYMM = _dYY * 100 + _dMM;     //基準年月
                int _DataYYMM;
                string fileYYMM;

                // 設定月数分経過した過去画像を削除する
                // ①スタッフ
                foreach (string files in System.IO.Directory.GetFiles(global.sTIF, "*.tif"))
                {
                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }

                // ②パートタイマー
                foreach (string files in System.IO.Directory.GetFiles(global.sTIF2, "*.tif"))
                {
                    //ファイル名より年月を取得する
                    fileYYMM = System.IO.Path.GetFileName(files).Substring(0, 6);

                    if (Utility.NumericCheck(fileYYMM))
                    {
                        _DataYYMM = int.Parse(fileYYMM);

                        //基準年月以前なら削除する
                        if (_DataYYMM <= _dYYMM) File.Delete(files);
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("過去画像削除中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
                return;
            }
        }

        /// <summary>
        /// 設定月数を経過したバックアップファイルを削除します
        /// </summary>
        /// <param name="sSpan">経過月数</param>
        /// <param name="sDir">ディレクトリ</param>
        private void delBackUpFiles(int sSpan, string sDir)
        {
            if (System.IO.Directory.Exists(sDir))
            {
                if (sSpan > 0)
                {
                    foreach (string fName in System.IO.Directory.GetFiles(sDir, "*.tif"))
                    {
                        string f = System.IO.Path.GetFileName(fName);

                        // ファイル名の長さを検証（日付情報あり？）
                        if (f.Length > 12)
                        {
                            // ファイル名から日付部分を取得します
                            DateTime dt;
                            string stDt = f.Substring(0, 4) + "/" + f.Substring(4, 2) + "/" + f.Substring(6, 2);
                            if (DateTime.TryParse(stDt, out dt))
                            {
                                // 設定月数を加算した日付を取得します
                                DateTime Fdt = dt.AddMonths(sSpan);

                                // 今日の日付と比較して設定月数を加算したファイル日付が既に経過している場合、ファイルを削除します
                                if (DateTime.Today.CompareTo(Fdt) == 1)
                                {
                                    System.IO.File.Delete(fName);
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// MDB明細データの件数をカウントする
        /// </summary>
        /// <returns>レコード件数</returns>
        private int CountMDBitem()
        {
            int rCnt = 0;

            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();
            OleDbDataReader dR;

            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("SELECT 勤務票ヘッダ.*,勤務票明細.* from ");
            sb.Append("勤務票ヘッダ inner join 勤務票明細 ");
            sb.Append("on 勤務票ヘッダ.ID = 勤務票明細.ヘッダID ");
            sb.Append("where データ区分=? ");
            sb.Append("order by 勤務票ヘッダ.ID,勤務票明細.ヘッダID ");
            sCom.CommandText = sb.ToString();
            sCom.Parameters.Clear();
            //sCom.Parameters.AddWithValue("@year", global.sYear);
            //sCom.Parameters.AddWithValue("@month", global.sMonth);
            sCom.Parameters.AddWithValue("@kbn", _usrSel);
            dR = sCom.ExecuteReader();

            while (dR.Read())
            {
                //データ件数加算
                rCnt++;
            }

            dR.Close();
            sCom.Connection.Close();

            return rCnt;
        }

        private void frmCorrect_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Tag.ToString() != END_MAKEDATA)
            {
                if (MessageBox.Show("終了します。よろしいですか", "終了確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    e.Cancel = true;
                    return;
                }

                // カレントデータ更新
                CurDataUpDate(_cI);
            }

            // 解放する
            this.Dispose();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            //フォームを閉じる
            this.Tag = END_BUTTON;
            this.Close();
        }

        /// <summary>
        /// MDBファイルを最適化する
        /// </summary>
        private void mdbCompact()
        {
            try
            {
                JRO.JetEngine jro = new JRO.JetEngine();

                string OldDb = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBFILENAME;

                string NewDb = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                    "Data Source=" + Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBTEMPFILE;

                // 最適化した一時MDBを作成する
                jro.CompactDatabase(OldDb, NewDb);

                //今までのバックアップファイルを削除する
                System.IO.File.Delete(Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBBACKUP);

                //今までのファイルをバックアップとする
                System.IO.File.Move(Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBFILENAME,
                                    Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBBACKUP);

                //一時ファイルをMDBファイルとする
                System.IO.File.Move(Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBTEMPFILE,
                                    Properties.Settings.Default.instPath + Properties.Settings.Default.MDB + global.MDBFILENAME);
            }
            catch (Exception e)
            {
                MessageBox.Show("MDB最適化中" + Environment.NewLine + e.Message, "エラー", MessageBoxButtons.OK);
            }
        }

        private void txtNo_Leave(object sender, EventArgs e)
        {
            txtNo.Text = txtNo.Text.PadLeft(7, '0');
            GetShoteiData(txtNo.Text);
        }

        private void txtNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                frmStaffSelect frmS = new frmStaffSelect(_usrSel);
                frmS.ShowDialog();
                if (frmS._msCode != string.Empty)
                {
                    txtNo.Text = frmS._msCode;
                    GetShoteiData(txtNo.Text);
                }
                frmS.Dispose();
            }
        }

        /// <summary>
        /// 勤務票ヘッダデータと勤務表明細データに新規レコードを追加する
        /// </summary>
        private void AddNewData(int usr)
        {
            // IDを取得します
            string _ID = string.Format("{0:0000}", DateTime.Today.Year) + 
                         string.Format("{0:00}", DateTime.Today.Month) + 
                         string.Format("{0:00}", DateTime.Today.Day) + 
                         string.Format("{0:00}", DateTime.Now.Hour) + 
                         string.Format("{0:00}", DateTime.Now.Minute) + 
                         string.Format("{0:00}", DateTime.Now.Second) + "001";

            // データベースへ接続します
            SysControl.SetDBConnect Con = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = Con.cnOpen();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                // 勤務票ヘッダテーブル追加登録
                StringBuilder sb = new StringBuilder();
                sb.Clear();
                sb.Append("insert into 勤務票ヘッダ ");
                sb.Append("(ID,年,月,データ区分) values (?,?,?,?)");

                sCom.CommandText = sb.ToString();
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@id", _ID);
                sCom.Parameters.AddWithValue("@year", global.sYear);
                sCom.Parameters.AddWithValue("@month", global.sMonth);
                sCom.Parameters.AddWithValue("@kbn", usr);
                sCom.ExecuteNonQuery();

                // 勤務票明細テーブル追加登録
                sb.Clear();
                sb.Append("insert into 勤務票明細 ");
                sb.Append("(ヘッダID, 日付, 休日) values (?,?,?)");
                sCom.CommandText = sb.ToString();

                int sDays = 1;
                DateTime dt;

                string tempDt = global.sYear + "/" + global.sMonth + "/" + sDays.ToString();

                // 存在する日付のときにMDBへ登録する
                while (DateTime.TryParse(tempDt, out dt))
                {
                    sCom.Parameters.Clear();

                    // ヘッダID
                    sCom.Parameters.AddWithValue("@ID", _ID);

                    // 日付
                    sCom.Parameters.AddWithValue("@Days", sDays);

                    // 休日区分の設定
                    // ①曜日で判断します
                    string Youbi = ("日月火水木金土").Substring(int.Parse(dt.DayOfWeek.ToString("d")), 1);
                    int sHol = 0;
                    if (Youbi == "土")
                    {
                        sHol = global.hSATURDAY;
                    }
                    else if (Youbi == "日")
                    {
                        sHol = global.hHOLIDAY;
                    }
                    else
                    {
                        sHol = global.hWEEKDAY;
                    }

                    // ②休日テーブルを参照し休日に該当するか調べます
                    SysControl.SetDBConnect dc = new SysControl.SetDBConnect();
                    OleDbCommand sCom2 = new OleDbCommand();
                    sCom2.Connection = dc.cnOpen();
                    OleDbDataReader dr = null;
                    sCom2.CommandText = "select * from 休日 where 年=? and 月=? and 日=?";
                    sCom2.Parameters.Clear();
                    sCom2.Parameters.AddWithValue("@year", global.sYear);
                    sCom2.Parameters.AddWithValue("@Month", global.sMonth);
                    sCom2.Parameters.AddWithValue("@day", sDays);
                    dr = sCom2.ExecuteReader();
                    if (dr.Read())
                    {
                        sHol = global.hHOLIDAY;
                    }
                    dr.Close();
                    sCom2.Connection.Close();

                    sCom.Parameters.AddWithValue("@Hol", sHol);

                    // テーブル書き込みます
                    sCom.ExecuteNonQuery();

                    // 日付をインクリメントします
                    sDays++;
                    tempDt = global.sYear + "/" + global.sMonth + "/" + sDays.ToString();
                }

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤務票データ新規登録処理", MessageBoxButtons.OK);
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中の勤務票データを削除します。よろしいですか", "削除確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                return;

            // レコードを削除する
            DataDelete(_cI);

            //画像ファイルを削除する
            if (System.IO.File.Exists(_InPath + global.pblImageFile))
            {
                System.IO.File.Delete(_InPath + global.pblImageFile);
            }

            //テーブル件数カウント：ゼロならばプログラム終了
            if (CountMDB(_usrSel) == 0)
            {
                MessageBox.Show("全ての勤務票データが削除されました。処理を終了します。", "勤務票削除", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                //終了処理
                Environment.Exit(0);
            }

            //テーブルデータキー項目読み込み
            inData = LoadMdbID(_usrSel);

            //エラー情報初期化
            ErrInitial();

            //レコードを表示
            if (inData.Length - 1 < _cI) _cI = inData.Length - 1;

            DataShow(_cI, inData, this.dg1);
        }

        /// <summary>
        /// カレント勤務票データを削除します
        /// </summary>
        /// <param name="iX">勤務票データインデックス</param>
        private void DataDelete(int iX)
        {
            //カレントデータを削除します
            //MDB接続
            SysControl.SetDBConnect dCon = new SysControl.SetDBConnect();
            OleDbCommand sCom = new OleDbCommand();
            sCom.Connection = dCon.cnOpen();

            // 画像ファイル名を取得します
            string sImgNm = string.Empty;
            OleDbDataReader dR;
            sCom.CommandText = "select 画像名 from 勤務票ヘッダ where ID = ?";
            sCom.Parameters.Clear();
            sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);
            dR = sCom.ExecuteReader();
            while (dR.Read())
            {
                sImgNm = dR["画像名"].ToString();
            }
            dR.Close();

            //トランザクション開始
            OleDbTransaction sTran = null;
            sTran = sCom.Connection.BeginTransaction();
            sCom.Transaction = sTran;

            try
            {
                //勤務票ヘッダデータを削除します
                sCom.CommandText = "delete from 勤務票ヘッダ where ID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);
                sCom.ExecuteNonQuery();

                //勤務票明細データを削除します
                sCom.CommandText = "delete from 勤務票明細 where ヘッダID = ?";
                sCom.Parameters.Clear();
                sCom.Parameters.AddWithValue("@ID", inData[iX]._sID);
                sCom.ExecuteNonQuery();

                //画像ファイルを削除する
                if (System.IO.File.Exists(_InPath + sImgNm))
                {
                    System.IO.File.Delete(_InPath + sImgNm);
                }

                // トランザクションコミット
                sTran.Commit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("勤務表の削除に失敗しました" + Environment.NewLine + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                // トランザクションロールバック
                sTran.Rollback();
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open) sCom.Connection.Close();
            }
        }

        private void chkRidatsu_CheckedChanged(object sender, EventArgs e)
        {
            if (chkRidatsu.Checked == true) chkRidatsu.ForeColor = Color.Red;
            else chkRidatsu.ForeColor = Color.Black;
        }

    }
}
