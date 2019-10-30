using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HSS.Model;

namespace HSS.ScanOcr
{
    public partial class frmPcStaffSelect : Form
    {
        public int _pcSel { get; set; }
        public int _usrSel { get; set; }

        public frmPcStaffSelect()
        {
            InitializeComponent();
        }

        private void frmPcStaffSelect_Load(object sender, EventArgs e)
        {
            // フォームの最大サイズ、最小サイズの設定
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // PC1を既定値とする
            this.rPcBtn1.Checked = true;

            // スタッフを既定値とする
            this.rStaffBtn1.Checked = true;

            // 値選択既定値
            _pcSel = global.END_SELECT;
            _usrSel = global.END_SELECT;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            // PC選択
            if (rPcBtn1.Checked) _pcSel = global.PC1SELECT;
            else if (rPcBtn2.Checked) _pcSel = global.PC2SELECT;

            // スタッフ選択
            if (rStaffBtn1.Checked) _usrSel = global.STAFF_SELECT;
            else if (rStaffBtn2.Checked) _usrSel = global.PART_SELECT;

            this.Close();
        }

        private void btnNo_Click(object sender, EventArgs e)
        {
            // 終了
            this.Close();
        }
    }
}
