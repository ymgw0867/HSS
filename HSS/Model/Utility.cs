using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Text.RegularExpressions;

namespace HSS.Model
{
    class Utility
    {
        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">Height</param>
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// ウィンドウ最小サイズの設定
        /// </summary>
        /// <param name="tempFrm">対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">width</param>
        /// <param name="hSize">height</param>
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// 文字列の値が数字かチェックする
        /// </summary>
        /// <param name="tempStr">検証する文字列</param>
        /// <returns>数字:true,数字でない:false</returns>
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }

        //処理モード
        public class formMode
        {
            public int Mode { get; set; }
            public int ID { get; set; }
        }

        //選択された会社のIDコード
        public class SelectCompany
        {
            public int ID { get; set; }
            public string Name { get; set; }
        }

        /// <summary>
        /// 郵便番号の正規表現チェック
        /// </summary>
        /// <param name="tempValue">対象となる郵便番号(文字列）</param>
        /// <returns>正しい：true,　正しくない：False</returns>
        public static Boolean RegexZip(string tempValue)
        {
            if (Regex.IsMatch(tempValue, @"^\d{3}-\d{4}$") == false)
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 国内電話番号の正規表現チェック
        /// </summary>
        /// <param name="tempValue">対象となる電話番号(文字列）</param>
        /// <returns>正しい：true,　正しくない：False</returns>
        public static Boolean RegexTel(string tempValue)
        {
            if (Regex.IsMatch(tempValue, @"^\d{2}-\d{4}-\d{4}$") == false)
            {
                if (Regex.IsMatch(tempValue, @"^\d{3}-\d{3}-\d{4}$") == false)
                {
                    if (Regex.IsMatch(tempValue, @"^\d{4}-\d{2}-\d{4}$") == false)
                    {
                        if (Regex.IsMatch(tempValue, @"^\d{5}-\d{1}-\d{4}$") == false)
                        {
                            return false;
                        }
                    }
                }
            }

            return true;

            //Regex reg = new Regex(@"^\d{3}-\d{4}$");

            //if (reg.IsMatch(this.txtZipCode.Text) == false)
            //{
            //    MessageBox.Show("郵便番号の書式（999-9999）が正しくありません", "郵便番号書式エラー", MessageBoxButton.OK);
            //    this.txtZipCode.Focus();
            //    return false;
            //}

        }

        /// <summary>
        /// 文字列を指定の文字数分取得する
        /// </summary>
        /// <param name="stemp">対象の文字列</param>
        /// <param name="sLen">文字数</param>
        /// <returns></returns>
        public static string GetSubString(string stemp, int sLen)
        {
            //nullまたはemptyのときはstring.emptyを返す
            if (stemp == null || stemp == string.Empty) return string.Empty;
            
            //対象文字列が指定文字数未満のときはそのまま返す
            if (stemp.Length <= sLen) return stemp;
            
            //最初の文字から指定文字数分取得して返す
            return stemp.Trim().Substring(0, sLen);
        }

        /// <summary>
        /// 定義済みの色の名前（文字列）からそのColor型を作成する
        /// </summary>
        /// <param name="c">色の名前（文字列）</param>
        /// <returns>color型インスタンス</returns>
        public static Color SetBackcolor(string c)
        {
            Color cl;
            try
            {
                cl = Color.FromName(c);
            }
            catch (Exception)
            {
                cl = Color.FromName("White");
            }

            return cl;
        }

        /// <summary>
        /// チェックデジット（JAN）を求める
        /// </summary>
        /// <param name="sMod">モジュラス値(10)</param>
        /// <param name="sWait">ウェイト値(3)</param>
        /// <param name="sCol">コード桁数</param>
        /// <param name="sVal">対象となるコード</param>
        /// <returns></returns>
        public static int GetjanCdigit(int sMod, int sWait, int sCol, string sVal)
        {
            int sumKi = 0;
            int sumGu = 0;
            int Digit;

            //奇数位置の値合計
            for (int i = sCol - 1; i >= 0; i -= 2)
                sumKi += int.Parse(sVal.Substring(i, 1));

            //ウェイトを乗算
            sumKi *= sWait;

            //偶数位置の値合計
            for (int i = sCol - 2; i >= 0; i -= 2)
                sumGu += int.Parse(sVal.Substring(i, 1));

            //奇数と偶数の総和を求める
            int sumTL = sumKi + sumGu;

            //モジュラス値(10)から１の位を差し引く
            int sumPl = sumTL % 10;             //１の位の値を取得
            if (sumPl == 0) Digit = 0;      //チェックデジット値は0
            else Digit = sMod - sumPl;      //モジュラス値(10)から差引いた値がチェックデジット

            return Digit;
        }

        /// <summary>
        /// コンボボックスクラス
        /// </summary>
        public class comboBkDel
        {
            public int ID { get; set; }
            public string Name { get; set; }

            /// <summary>
            /// コンボボックスデータロード
            /// </summary>
            /// <param name="tempBox">ロード先コンボボックスオブジェクト名</param>
            public static void Load(ComboBox tempBox)
            {
                try
                {
                    comboBkDel cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "ID";

                    for (int i = 0; i <= Properties.Settings.Default.BkdelMax; i++)
                    {
                        cmb1 = new comboBkDel();
                        cmb1.ID = i;
                        cmb1.Name = i.ToString();
                        tempBox.Items.Add(cmb1);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "コンボボックスロード");
                }
            }

            /// <summary>
            /// コンボ表示
            /// </summary>
            /// <param name="tempBox">コンボボックスオブジェクト</param>
            /// <param name="id">ID</param>
            public static int selectedIndex(ComboBox tempBox, int id)
            {
                comboBkDel cmbS = new comboBkDel();
                int reIndex = -1;

                for (int iX = 0; iX < tempBox.Items.Count; iX++)
                {
                    tempBox.SelectedIndex = iX;
                    cmbS = (comboBkDel)tempBox.SelectedItem;

                    if (cmbS.ID == id)
                    {
                        reIndex = iX;
                        break;
                    }
                }

                return reIndex;
            }
        }

        /// <summary>
        /// 文字列を指定文字数をＭＡＸとして返します
        /// </summary>
        /// <param name="s">文字列</param>
        /// <param name="n">文字数</param>
        /// <returns>文字数範囲内の文字列</returns>
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;
            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        /// <summary>
        /// emptyを"0"に置き換える
        /// </summary>
        /// <param name="tempStr">stringオブジェクト</param>
        /// <returns>nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        /// <summary>
        /// 文字型をIntへ変換して返す（数値でないときは０を返す）
        /// </summary>
        /// <param name="tempStr">文字型の値</param>
        /// <returns>Int型の値</returns>
        public static int StrToInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }

        /// <summary>
        /// Nullをstring.Empty("")に置き換える
        /// </summary>
        /// <param name="tempStr">stringオブジェクト</param>
        /// <returns>nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        /// <summary>
        /// Nullをstring.Empty("")に置き換える
        /// </summary>
        /// <param name="tempStr">stringオブジェクト</param>
        /// <returns>nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// <summary>
        /// 時分表示を分数に変換した値を返す
        /// </summary>
        /// <param name="sTime">時分表示文字列(hh:mm, hhmm)</param>
        /// <returns>分数表示値</returns>
        public static int TimeToMin(string sTime)
        {
            int hh = 0;
            int mm = 0;

            if (sTime.Length == 5)
            {
                hh = int.Parse(sTime.Substring(0, 2)) * 60;
                mm = int.Parse(sTime.Substring(3, 2));
            }
            else if (sTime.Length == 4)
            {
                hh = int.Parse(sTime.Substring(0, 2)) * 60;
                mm = int.Parse(sTime.Substring(2, 2));
            }

            return hh + mm;
        }

        /// <summary>
        /// 経過時間を返す
        /// </summary>
        /// <param name="s">開始時間</param>
        /// <param name="e">終了時間</param>
        /// <returns>経過時間</returns>
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        public static double fncTime10(double sMytime)
        {
            int Seisu = 0;
            int Shosu = 0;

            double s = sMytime / 60;
            string[] iTime = s.ToString().Split('.');
            Seisu = int.Parse(iTime[0]);

            if (iTime.Length == 1)
                return s;
            else 
            {
                if (iTime[1].Length > 1)
                {
                    if (iTime[1].Substring(0, 1) == "9")
                    {
                        Seisu = Seisu + 1;
                        Shosu = 0;
                    }
                    else
                    {
                        Shosu = int.Parse(iTime[1].Substring(0, 1));
                        Shosu = Shosu + 1;
                    }

                    s = double.Parse(Seisu.ToString() + "." + Shosu.ToString());
                }
            }

            return s;
        }

        /// <summary>
        /// 小数点を消去して小数点二位以下は切り捨てる
        /// 少数点以下がないときは０を置く
        /// 0のときは小数点以下はなし
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string fncTen3(double s)
        {
            if (s == 0) return "0";

            string rtn = string.Empty;
            string[] iTime = s.ToString().Split('.');

            if (iTime.Length == 1) rtn = iTime[0] + "0";
            else rtn = iTime[0] + iTime[1].Substring(0, 1);

            return rtn;
        }

        /// <summary>
        /// 小数点以下は一位で切り捨てる
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string fncTen2(double s)
        {
            string rtn = string.Empty;
            string[] iTime = s.ToString().Split('.');

            if (iTime.Length == 1)
                rtn = iTime[0];
            else
            {
                rtn = iTime[0] + "." + iTime[1].Substring(0, 1);
            }

            return rtn;
        }

        /// <summary>
        /// 小数点以下が０のときカットする
        /// </summary>
        /// <param name="s">対象文字列</param>
        /// <returns>戻り値</returns>
        public static string fncShosuZeroCut(string s)
        {
            string rtn = string.Empty;
            string[] iTime = s.ToString().Split('.');

            if (iTime[1] == "0") rtn = iTime[0];
            else rtn = s;

            return rtn;
        }
    }
}
