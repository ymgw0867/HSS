using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HSS.Model;

namespace HSS.Combi
{
    public partial class frmCsvCombi : Form
    {
        public frmCsvCombi()
        {
            InitializeComponent();
        }

        // 初期ロードステータス
        private bool _load = false;

        // スタッフデータのファイルタイプ
        int cVal = 0;   // 0:通常、1:離脱

        // 読み込みCSV
        private string[] sBuf;

        private void frmCsvCombi_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            GridViewSetting(dataGridView1);
            _load = true;

            // スタッフデータを初期選択とする
            radioButton1.Checked = true;
            loadData(dataGridView1, global.sDAT, "*.csv");
        }

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
                tempDGV.Columns.Add("c1", "ファイル名");

                tempDGV.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = true;

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

        private void loadData(DataGridView d, string sPath, string pat)
        {
            int i = 0;
            d.RowCount = 0;

            foreach (var f in System.IO.Directory.GetFiles(sPath, pat))
            {
                d.Rows.Add();
                d[0, i].Value = System.IO.Path.GetFileName(f);
                i++;
            }

            d.CurrentCell = null;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (_load)
            {
                if (radioButton1.Checked) loadData(dataGridView1, global.sDAT, "*.csv");
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (_load)
            {
                if (radioButton2.Checked) loadData(dataGridView1, global.sDAT2, "*.*");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmCsvCombi_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fName = string.Empty;

            // 指定ファイルがあるか
            if (dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("結合するファイルを選択してください。", "ファイル結合", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 通常データと離脱データの混在指定は不可
            if (!combineFileType())
            {
                MessageBox.Show("通常データと離脱データを混在して結合することは出来ません。", "ファイル結合", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("選択されたデータファイルを結合します。よろしいですか？", "ファイル結合",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            // ダイアログボックス表示
            saveFileDialog1.Filter = "データファイル(*.csv;*.txt;*.dat)|*.csv;*.txt;*.dat";
            saveFileDialog1.Title = "結合したファイルの保存場所とファイル名を指定してください";
                
            if (radioButton1.Checked)
            {
                saveFileDialog1.InitialDirectory = global.sDAT;
                saveFileDialog1.FileName = "スタッフ勤務報告" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";
            }
            else
            {
                saveFileDialog1.InitialDirectory = global.sDAT2;
                saveFileDialog1.FileName = "パート勤務報告" + DateTime.Now.ToString("yyyy年MM月dd日HH時mm分ss秒") + ".csv";
            }

            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // 出力ファイル名を取得します
                fName = saveFileDialog1.FileName;

                // 入力ファイルのパスを設定します
                string sPath = string.Empty;
                if (radioButton1.Checked) sPath = global.sDAT;
                else if (radioButton2.Checked) sPath = global.sDAT2;

                // ファイルの結合処理を行います
                CsvCombine(dataGridView1, sPath, fName);

            }
            else return;
        }

        /// <summary>
        /// 結合するデータが通常、離脱混在していないか調査する
        /// </summary>
        /// <returns>True:混在していない、false:混在している</returns>
        private bool combineFileType()
        {
            bool rtn = true;
            int gVal = 0;

            for (int i = 0; i < dataGridView1.SelectedRows.Count; i++)
            {
                int r = dataGridView1.SelectedRows[i].Index;
                string h = dataGridView1[0, r].Value.ToString();

                // 通常データか離脱データか判断します
                if (h.Substring(4,1) == "R") gVal = 1;
                else gVal = 0;

                if (i == 0) cVal = gVal;
                else
                {
                    if (cVal != gVal)　// 照合します
                    {
                        rtn = false;
                        break;
                    }
                }                
            }
            return rtn;
        }

        private void CsvCombine(DataGridView d, string inPath, string outFile)
        {
            int rCnt = 1;

            //sBuf = new string[d.SelectedRows.Count];
            sBuf = new string[1];

            // 出力ファイルインスタンス作成
            System.IO.StreamWriter cFile = new System.IO.StreamWriter(outFile, true, System.Text.Encoding.GetEncoding(932));

            // 出力ファイルに書き出します
            for (int i = 0; i < d.SelectedRows.Count; i++)
            {
                // 入力ファイルのパスを取得します
                string sPath = inPath + d[0, d.SelectedRows[i].Index].Value.ToString();

                // 入力CSVファイルの全ての行を読み取ります
                var s = System.IO.File.ReadAllLines(sPath, Encoding.Default);
                foreach (var stBuffer in s)
                {
                    //１行ずつ出力ファイルに書き出します
                    //sBuf[i] = stBuffer;

                    //2件目以降は要素数を追加
                    if (rCnt > 1) sBuf.CopyTo(sBuf = new string[rCnt], 0);
                    sBuf[rCnt - 1] = stBuffer;

                    rCnt++;
                }
            }

            // 並べなおします
            string nowTenpo = string.Empty;
            string nxtTenpo = string.Empty;

            for (int sCnt = sBuf.Length - 1; sCnt >= 0; sCnt--)
            {
                for (int lCnt = 0; lCnt < sCnt; lCnt++)
                {
                    if (radioButton2.Checked || cVal == 1)　// パートタイマーまたはスタッフ離脱データのとき
                    {
                        nowTenpo = sBuf[lCnt].Substring(3, 7);
                        nxtTenpo = sBuf[lCnt + 1].Substring(3, 7);
                    }
                    else if (radioButton1.Checked && cVal == 0) // スタッフ通常データのとき
                    {
                        nowTenpo = sBuf[lCnt].Substring(0, 7);
                        nxtTenpo = sBuf[lCnt + 1].Substring(0, 7);
                    }

                    if (Utility.StrToInt(nowTenpo) > Utility.StrToInt(nxtTenpo))
                    {
                        string mv = sBuf[lCnt + 1];
                        sBuf[lCnt + 1] = sBuf[lCnt];
                        sBuf[lCnt] = mv;
                    }
                }
            }
            
            // 書き出します
            int iSQ = 1;
            for (int i = 0; i < sBuf.Length; i++)
			{
                if (radioButton2.Checked)
                {
                    //sBuf[i] = sBuf[i].Substring(0, 96) + iSQ.ToString().PadLeft(5, '0') + sBuf[i].Substring(101, 7);

                    string t = sBuf[i].Substring(94, 2).Trim().PadLeft(2, '0');
                    sBuf[i] = sBuf[i].Substring(0, 94) + t + iSQ.ToString().PadLeft(5, '0') + sBuf[i].Substring(101, 7);
                    iSQ++;
                }

                cFile.WriteLine(sBuf[i]);
			}

            cFile.Close();

            // 元データの削除
            string msg = "受け渡しデータを結合しました。元のデータファイルを削除しますか？" + Environment.NewLine + Environment.NewLine +
                         "削除する場合には「はい」を削除しない場合には「いいえ」を選択してください。";

            if (MessageBox.Show(msg,"ファイル結合",MessageBoxButtons.YesNo,MessageBoxIcon.Information,MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No) return;

            for (int i = 0; i < d.SelectedRows.Count; i++)
            {
                // ファイルを削除します
                string sPath = inPath + d[0, d.SelectedRows[i].Index].Value.ToString();
                System.IO.File.Delete(sPath);
            }

            MessageBox.Show("元のデータを削除しました");

            if (radioButton1.Checked) loadData(dataGridView1, global.sDAT, "*.csv");
            else if (radioButton2.Checked) loadData(dataGridView1, global.sDAT2, "*.*");
        }

        private void CsvCombineTest(DataGridView d, string inPath)
        {
            // 出力ファイルストリーム生成
            System.IO.FileStream fsOut = new System.IO.FileStream("c:\\ketsugocsv.csv",System.IO.FileMode.Append,System.IO.FileAccess.Write);

            // 入力ファイル
            for (int i = 0; i < d.SelectedRows.Count; i++)
			{
                string sPath = inPath + d[0, d.SelectedRows[i].Index].Value.ToString();
                System.IO.FileStream s = new System.IO.FileStream(sPath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                int t;
                while ((t = s.ReadByte()) != -1)
                {
                    fsOut.WriteByte((byte)t);
                }
                s.Close();
			}
            fsOut.Close();
        }
    }
}
