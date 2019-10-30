using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HSS.Model;

namespace HSS.PrePrint
{
    public partial class frmUserSelect : Form
    {
        public int _usrSel { get; set; }

        public frmUserSelect()
        {
            InitializeComponent();
        }

        private void frmUserSelect_Load(object sender, EventArgs e)
        {
            // フォームの最大サイズ、最小サイズの設定
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // スタッフを既定とする
            rBtn1.Checked = true;

            _usrSel = global.END_SELECT;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (rBtn1.Checked) _usrSel = global.STAFF_SELECT;
            else _usrSel = global.PART_SELECT;

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _usrSel = global.END_SELECT;
            this.Close();
        }

        private void frmUserSelect_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }
    }
}
