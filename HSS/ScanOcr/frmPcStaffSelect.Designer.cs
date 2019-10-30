namespace HSS.ScanOcr
{
    partial class frmPcStaffSelect
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPcStaffSelect));
            this.panel1 = new System.Windows.Forms.Panel();
            this.rPcBtn2 = new System.Windows.Forms.RadioButton();
            this.rPcBtn1 = new System.Windows.Forms.RadioButton();
            this.panel2 = new System.Windows.Forms.Panel();
            this.rStaffBtn2 = new System.Windows.Forms.RadioButton();
            this.rStaffBtn1 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnNo = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.rPcBtn2);
            this.panel1.Controls.Add(this.rPcBtn1);
            this.panel1.Location = new System.Drawing.Point(24, 39);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(302, 54);
            this.panel1.TabIndex = 0;
            // 
            // rPcBtn2
            // 
            this.rPcBtn2.AutoSize = true;
            this.rPcBtn2.Location = new System.Drawing.Point(149, 14);
            this.rPcBtn2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rPcBtn2.Name = "rPcBtn2";
            this.rPcBtn2.Size = new System.Drawing.Size(48, 22);
            this.rPcBtn2.TabIndex = 1;
            this.rPcBtn2.TabStop = true;
            this.rPcBtn2.Text = "PC2";
            this.rPcBtn2.UseVisualStyleBackColor = true;
            // 
            // rPcBtn1
            // 
            this.rPcBtn1.AutoSize = true;
            this.rPcBtn1.Location = new System.Drawing.Point(39, 14);
            this.rPcBtn1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rPcBtn1.Name = "rPcBtn1";
            this.rPcBtn1.Size = new System.Drawing.Size(48, 22);
            this.rPcBtn1.TabIndex = 0;
            this.rPcBtn1.TabStop = true;
            this.rPcBtn1.Text = "PC1";
            this.rPcBtn1.UseVisualStyleBackColor = true;
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.rStaffBtn2);
            this.panel2.Controls.Add(this.rStaffBtn1);
            this.panel2.Location = new System.Drawing.Point(24, 131);
            this.panel2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(302, 54);
            this.panel2.TabIndex = 1;
            // 
            // rStaffBtn2
            // 
            this.rStaffBtn2.AutoSize = true;
            this.rStaffBtn2.Location = new System.Drawing.Point(149, 14);
            this.rStaffBtn2.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rStaffBtn2.Name = "rStaffBtn2";
            this.rStaffBtn2.Size = new System.Drawing.Size(110, 22);
            this.rStaffBtn2.TabIndex = 1;
            this.rStaffBtn2.TabStop = true;
            this.rStaffBtn2.Text = "パートタイマー";
            this.rStaffBtn2.UseVisualStyleBackColor = true;
            // 
            // rStaffBtn1
            // 
            this.rStaffBtn1.AutoSize = true;
            this.rStaffBtn1.Location = new System.Drawing.Point(39, 14);
            this.rStaffBtn1.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.rStaffBtn1.Name = "rStaffBtn1";
            this.rStaffBtn1.Size = new System.Drawing.Size(74, 22);
            this.rStaffBtn1.TabIndex = 0;
            this.rStaffBtn1.TabStop = true;
            this.rStaffBtn1.Text = "スタッフ";
            this.rStaffBtn1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(27, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "修正作業を行うＰＣを選択してください";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(27, 111);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(260, 18);
            this.label2.TabIndex = 3;
            this.label2.Text = "処理を行うスタッフの種類を選択してください";
            this.label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(24, 209);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(148, 33);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "ＯＫ(&O)";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnNo
            // 
            this.btnNo.Location = new System.Drawing.Point(178, 209);
            this.btnNo.Name = "btnNo";
            this.btnNo.Size = new System.Drawing.Size(148, 33);
            this.btnNo.TabIndex = 5;
            this.btnNo.Text = "キャンセル(&Q)";
            this.btnNo.UseVisualStyleBackColor = true;
            this.btnNo.Click += new System.EventHandler(this.btnNo_Click);
            // 
            // frmPcStaffSelect
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(351, 253);
            this.Controls.Add(this.btnNo);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmPcStaffSelect";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "スキャンOCR処理選択";
            this.Load += new System.EventHandler(this.frmPcStaffSelect_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton rPcBtn2;
        private System.Windows.Forms.RadioButton rPcBtn1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.RadioButton rStaffBtn2;
        private System.Windows.Forms.RadioButton rStaffBtn1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnNo;
    }
}