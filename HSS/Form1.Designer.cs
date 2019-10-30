namespace HSS
{
    partial class Form1
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.btnPrePrint = new System.Windows.Forms.Button();
            this.btnOCR = new System.Windows.Forms.Button();
            this.btnDataEntry = new System.Windows.Forms.Button();
            this.btnDataAddNew = new System.Windows.Forms.Button();
            this.btnDataCombi = new System.Windows.Forms.Button();
            this.btnSetup = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Window;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(32, 14);
            this.panel1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(347, 73);
            this.panel1.TabIndex = 0;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaption;
            this.label2.Location = new System.Drawing.Point(13, 33);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(317, 28);
            this.label2.TabIndex = 1;
            this.label2.Text = "勤怠データ ＯＣＲ接続プログラム";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("メイリオ", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(16, 7);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(311, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "八十二スタッフサービス株式会社";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnPrePrint
            // 
            this.btnPrePrint.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnPrePrint.Location = new System.Drawing.Point(32, 107);
            this.btnPrePrint.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnPrePrint.Name = "btnPrePrint";
            this.btnPrePrint.Size = new System.Drawing.Size(347, 36);
            this.btnPrePrint.TabIndex = 1;
            this.btnPrePrint.Text = "勤務票のプレ印刷(&E)";
            this.btnPrePrint.UseVisualStyleBackColor = true;
            this.btnPrePrint.Click += new System.EventHandler(this.btnPrePrint_Click);
            // 
            // btnOCR
            // 
            this.btnOCR.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnOCR.Location = new System.Drawing.Point(32, 153);
            this.btnOCR.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnOCR.Name = "btnOCR";
            this.btnOCR.Size = new System.Drawing.Size(347, 36);
            this.btnOCR.TabIndex = 2;
            this.btnOCR.Text = "ＯＣＲ認識処理(&R)";
            this.btnOCR.UseVisualStyleBackColor = true;
            this.btnOCR.Click += new System.EventHandler(this.btnOCR_Click);
            // 
            // btnDataEntry
            // 
            this.btnDataEntry.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDataEntry.Location = new System.Drawing.Point(32, 199);
            this.btnDataEntry.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnDataEntry.Name = "btnDataEntry";
            this.btnDataEntry.Size = new System.Drawing.Size(347, 36);
            this.btnDataEntry.TabIndex = 3;
            this.btnDataEntry.Text = "修正確認＆受け渡しデータ作成(&U)";
            this.btnDataEntry.UseVisualStyleBackColor = true;
            this.btnDataEntry.Click += new System.EventHandler(this.btnDataEntry_Click);
            // 
            // btnDataAddNew
            // 
            this.btnDataAddNew.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDataAddNew.Location = new System.Drawing.Point(32, 245);
            this.btnDataAddNew.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnDataAddNew.Name = "btnDataAddNew";
            this.btnDataAddNew.Size = new System.Drawing.Size(347, 36);
            this.btnDataAddNew.TabIndex = 4;
            this.btnDataAddNew.Text = "勤怠データ新規入力(&S)";
            this.btnDataAddNew.UseVisualStyleBackColor = true;
            this.btnDataAddNew.Click += new System.EventHandler(this.btnDataAddNew_Click);
            // 
            // btnDataCombi
            // 
            this.btnDataCombi.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnDataCombi.Location = new System.Drawing.Point(32, 291);
            this.btnDataCombi.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnDataCombi.Name = "btnDataCombi";
            this.btnDataCombi.Size = new System.Drawing.Size(347, 36);
            this.btnDataCombi.TabIndex = 5;
            this.btnDataCombi.Text = "受け渡しデータの結合(&K)";
            this.btnDataCombi.UseVisualStyleBackColor = true;
            this.btnDataCombi.Click += new System.EventHandler(this.btnDataCombi_Click);
            // 
            // btnSetup
            // 
            this.btnSetup.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnSetup.Location = new System.Drawing.Point(32, 337);
            this.btnSetup.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnSetup.Name = "btnSetup";
            this.btnSetup.Size = new System.Drawing.Size(347, 36);
            this.btnSetup.TabIndex = 6;
            this.btnSetup.Text = "環境設定(&C)";
            this.btnSetup.UseVisualStyleBackColor = true;
            this.btnSetup.Click += new System.EventHandler(this.btnSetup_Click);
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnExit.Location = new System.Drawing.Point(256, 425);
            this.btnExit.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(123, 36);
            this.btnExit.TabIndex = 8;
            this.btnExit.Text = "終了(&Q)";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("メイリオ", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.button1.Location = new System.Drawing.Point(32, 381);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(347, 36);
            this.button1.TabIndex = 7;
            this.button1.Text = "休日登録(&H)";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(411, 472);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSetup);
            this.Controls.Add(this.btnDataCombi);
            this.Controls.Add(this.btnDataAddNew);
            this.Controls.Add(this.btnDataEntry);
            this.Controls.Add(this.btnOCR);
            this.Controls.Add(this.btnPrePrint);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "メインメニュー";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnPrePrint;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnOCR;
        private System.Windows.Forms.Button btnDataEntry;
        private System.Windows.Forms.Button btnDataAddNew;
        private System.Windows.Forms.Button btnDataCombi;
        private System.Windows.Forms.Button btnSetup;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button button1;
    }
}

