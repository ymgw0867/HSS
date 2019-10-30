namespace HSS.Config
{
    partial class frmConfig
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmConfig));
            this.label1 = new System.Windows.Forms.Label();
            this.txtYear = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMonth = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtScanner = new System.Windows.Forms.TextBox();
            this.btnScanner = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.cmbBkdels = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.btnTif = new System.Windows.Forms.Button();
            this.txtTif = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.btnDat = new System.Windows.Forms.Button();
            this.txtDat = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.btnMsts = new System.Windows.Forms.Button();
            this.txtMsts = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.cmbBkdelp = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.BtnTif2 = new System.Windows.Forms.Button();
            this.txtTif2 = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.btnDat2 = new System.Windows.Forms.Button();
            this.txtDat2 = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.btnMstj = new System.Windows.Forms.Button();
            this.txtMstj = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnRtn = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.chkTimeYukyu = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label1.Location = new System.Drawing.Point(26, 37);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "処理年月：２０";
            // 
            // txtYear
            // 
            this.txtYear.Location = new System.Drawing.Point(121, 34);
            this.txtYear.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtYear.MaxLength = 2;
            this.txtYear.Name = "txtYear";
            this.txtYear.Size = new System.Drawing.Size(30, 27);
            this.txtYear.TabIndex = 0;
            this.txtYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtYear.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtYear.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            this.txtYear.Leave += new System.EventHandler(this.txtYear_Leave);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label2.Location = new System.Drawing.Point(152, 37);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(22, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "年";
            // 
            // txtMonth
            // 
            this.txtMonth.Location = new System.Drawing.Point(177, 34);
            this.txtMonth.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtMonth.MaxLength = 2;
            this.txtMonth.Name = "txtMonth";
            this.txtMonth.Size = new System.Drawing.Size(30, 27);
            this.txtMonth.TabIndex = 1;
            this.txtMonth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtMonth.Enter += new System.EventHandler(this.txtYear_Enter);
            this.txtMonth.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtYear_KeyPress);
            this.txtMonth.Leave += new System.EventHandler(this.txtYear_Leave);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label3.Location = new System.Drawing.Point(209, 37);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(22, 20);
            this.label3.TabIndex = 4;
            this.label3.Text = "月";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label4.Location = new System.Drawing.Point(252, 13);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(126, 20);
            this.label4.TabIndex = 5;
            this.label4.Text = "使用するスキャナ：";
            // 
            // txtScanner
            // 
            this.txtScanner.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtScanner.Location = new System.Drawing.Point(256, 30);
            this.txtScanner.Name = "txtScanner";
            this.txtScanner.Size = new System.Drawing.Size(377, 25);
            this.txtScanner.TabIndex = 3;
            this.txtScanner.Enter += new System.EventHandler(this.txtYear_Enter);
            // 
            // btnScanner
            // 
            this.btnScanner.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.btnScanner.Location = new System.Drawing.Point(256, 56);
            this.btnScanner.Name = "btnScanner";
            this.btnScanner.Size = new System.Drawing.Size(106, 30);
            this.btnScanner.TabIndex = 2;
            this.btnScanner.Text = "スキャナを選択";
            this.btnScanner.UseVisualStyleBackColor = true;
            this.btnScanner.Click += new System.EventHandler(this.btnScanner_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.cmbBkdels);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.btnTif);
            this.panel1.Controls.Add(this.txtTif);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.btnDat);
            this.panel1.Controls.Add(this.txtDat);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.btnMsts);
            this.panel1.Controls.Add(this.txtMsts);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Location = new System.Drawing.Point(21, 110);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(612, 232);
            this.panel1.TabIndex = 8;
            // 
            // cmbBkdels
            // 
            this.cmbBkdels.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbBkdels.FormattingEnabled = true;
            this.cmbBkdels.Location = new System.Drawing.Point(22, 188);
            this.cmbBkdels.Name = "cmbBkdels";
            this.cmbBkdels.Size = new System.Drawing.Size(49, 26);
            this.cmbBkdels.TabIndex = 6;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label9.Location = new System.Drawing.Point(19, 166);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(188, 18);
            this.label9.TabIndex = 19;
            this.label9.Text = "バックアップデータの自動削除：";
            // 
            // btnTif
            // 
            this.btnTif.Location = new System.Drawing.Point(546, 137);
            this.btnTif.Name = "btnTif";
            this.btnTif.Size = new System.Drawing.Size(52, 27);
            this.btnTif.TabIndex = 5;
            this.btnTif.Text = "参照";
            this.btnTif.UseVisualStyleBackColor = true;
            this.btnTif.Click += new System.EventHandler(this.btnTif_Click);
            // 
            // txtTif
            // 
            this.txtTif.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTif.Location = new System.Drawing.Point(22, 138);
            this.txtTif.Name = "txtTif";
            this.txtTif.Size = new System.Drawing.Size(523, 25);
            this.txtTif.TabIndex = 4;
            this.txtTif.Enter += new System.EventHandler(this.txtYear_Enter);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label8.Location = new System.Drawing.Point(19, 117);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(165, 18);
            this.label8.TabIndex = 16;
            this.label8.Text = "OCRデータバックアップ先：";
            // 
            // btnDat
            // 
            this.btnDat.Location = new System.Drawing.Point(546, 88);
            this.btnDat.Name = "btnDat";
            this.btnDat.Size = new System.Drawing.Size(52, 27);
            this.btnDat.TabIndex = 3;
            this.btnDat.Text = "参照";
            this.btnDat.UseVisualStyleBackColor = true;
            this.btnDat.Click += new System.EventHandler(this.btnDat_Click);
            // 
            // txtDat
            // 
            this.txtDat.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtDat.Location = new System.Drawing.Point(22, 89);
            this.txtDat.Name = "txtDat";
            this.txtDat.Size = new System.Drawing.Size(523, 25);
            this.txtDat.TabIndex = 2;
            this.txtDat.Enter += new System.EventHandler(this.txtYear_Enter);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label7.Location = new System.Drawing.Point(19, 68);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(140, 18);
            this.label7.TabIndex = 13;
            this.label7.Text = "受け渡しデータ作成先：";
            // 
            // btnMsts
            // 
            this.btnMsts.Location = new System.Drawing.Point(546, 38);
            this.btnMsts.Name = "btnMsts";
            this.btnMsts.Size = new System.Drawing.Size(52, 27);
            this.btnMsts.TabIndex = 1;
            this.btnMsts.Text = "参照";
            this.btnMsts.UseVisualStyleBackColor = true;
            this.btnMsts.Click += new System.EventHandler(this.btnMsts_Click);
            // 
            // txtMsts
            // 
            this.txtMsts.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMsts.Location = new System.Drawing.Point(22, 39);
            this.txtMsts.Name = "txtMsts";
            this.txtMsts.Size = new System.Drawing.Size(523, 25);
            this.txtMsts.TabIndex = 0;
            this.txtMsts.Enter += new System.EventHandler(this.txtYear_Enter);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label6.Location = new System.Drawing.Point(19, 18);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(212, 18);
            this.label6.TabIndex = 10;
            this.label6.Text = "スタッフ用マスターデータファイル：";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label10.Location = new System.Drawing.Point(70, 191);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(188, 18);
            this.label10.TabIndex = 21;
            this.label10.Text = "ヶ月経過したら自動的に削除する";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label5.Location = new System.Drawing.Point(27, 101);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 20);
            this.label5.TabIndex = 9;
            this.label5.Text = "スタッフ用設定";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label11.Location = new System.Drawing.Point(27, 356);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(87, 20);
            this.label11.TabIndex = 11;
            this.label11.Text = "パート用設定";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.cmbBkdelp);
            this.panel2.Controls.Add(this.label12);
            this.panel2.Controls.Add(this.BtnTif2);
            this.panel2.Controls.Add(this.txtTif2);
            this.panel2.Controls.Add(this.label13);
            this.panel2.Controls.Add(this.btnDat2);
            this.panel2.Controls.Add(this.txtDat2);
            this.panel2.Controls.Add(this.label14);
            this.panel2.Controls.Add(this.btnMstj);
            this.panel2.Controls.Add(this.txtMstj);
            this.panel2.Controls.Add(this.label15);
            this.panel2.Controls.Add(this.label16);
            this.panel2.Location = new System.Drawing.Point(21, 365);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(612, 232);
            this.panel2.TabIndex = 10;
            // 
            // cmbBkdelp
            // 
            this.cmbBkdelp.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.cmbBkdelp.FormattingEnabled = true;
            this.cmbBkdelp.Location = new System.Drawing.Point(22, 188);
            this.cmbBkdelp.Name = "cmbBkdelp";
            this.cmbBkdelp.Size = new System.Drawing.Size(49, 26);
            this.cmbBkdelp.TabIndex = 6;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label12.Location = new System.Drawing.Point(19, 166);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(188, 18);
            this.label12.TabIndex = 19;
            this.label12.Text = "バックアップデータの自動削除：";
            // 
            // BtnTif2
            // 
            this.BtnTif2.Location = new System.Drawing.Point(546, 137);
            this.BtnTif2.Name = "BtnTif2";
            this.BtnTif2.Size = new System.Drawing.Size(52, 27);
            this.BtnTif2.TabIndex = 5;
            this.BtnTif2.Text = "参照";
            this.BtnTif2.UseVisualStyleBackColor = true;
            this.BtnTif2.Click += new System.EventHandler(this.BtnTif2_Click);
            // 
            // txtTif2
            // 
            this.txtTif2.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtTif2.Location = new System.Drawing.Point(22, 138);
            this.txtTif2.Name = "txtTif2";
            this.txtTif2.Size = new System.Drawing.Size(523, 25);
            this.txtTif2.TabIndex = 4;
            this.txtTif2.Enter += new System.EventHandler(this.txtYear_Enter);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label13.Location = new System.Drawing.Point(19, 117);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(165, 18);
            this.label13.TabIndex = 16;
            this.label13.Text = "OCRデータバックアップ先：";
            // 
            // btnDat2
            // 
            this.btnDat2.Location = new System.Drawing.Point(546, 88);
            this.btnDat2.Name = "btnDat2";
            this.btnDat2.Size = new System.Drawing.Size(52, 27);
            this.btnDat2.TabIndex = 3;
            this.btnDat2.Text = "参照";
            this.btnDat2.UseVisualStyleBackColor = true;
            this.btnDat2.Click += new System.EventHandler(this.btnDat2_Click);
            // 
            // txtDat2
            // 
            this.txtDat2.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtDat2.Location = new System.Drawing.Point(22, 89);
            this.txtDat2.Name = "txtDat2";
            this.txtDat2.Size = new System.Drawing.Size(523, 25);
            this.txtDat2.TabIndex = 2;
            this.txtDat2.Enter += new System.EventHandler(this.txtYear_Enter);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label14.Location = new System.Drawing.Point(19, 68);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(140, 18);
            this.label14.TabIndex = 13;
            this.label14.Text = "受け渡しデータ作成先：";
            // 
            // btnMstj
            // 
            this.btnMstj.Location = new System.Drawing.Point(546, 38);
            this.btnMstj.Name = "btnMstj";
            this.btnMstj.Size = new System.Drawing.Size(52, 27);
            this.btnMstj.TabIndex = 1;
            this.btnMstj.Text = "参照";
            this.btnMstj.UseVisualStyleBackColor = true;
            this.btnMstj.Click += new System.EventHandler(this.btnMstj_Click);
            // 
            // txtMstj
            // 
            this.txtMstj.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.txtMstj.Location = new System.Drawing.Point(22, 39);
            this.txtMstj.Name = "txtMstj";
            this.txtMstj.Size = new System.Drawing.Size(523, 25);
            this.txtMstj.TabIndex = 0;
            this.txtMstj.Enter += new System.EventHandler(this.txtYear_Enter);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label15.Location = new System.Drawing.Point(19, 18);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(248, 18);
            this.label15.TabIndex = 10;
            this.label15.Text = "パートタイマー用マスターデータファイル：";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("メイリオ", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.label16.Location = new System.Drawing.Point(70, 191);
            this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(188, 18);
            this.label16.TabIndex = 21;
            this.label16.Text = "ヶ月経過したら自動的に削除する";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(282, 610);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(163, 32);
            this.btnSave.TabIndex = 4;
            this.btnSave.Text = "設定を保存(&O)";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnRtn
            // 
            this.btnRtn.Location = new System.Drawing.Point(550, 611);
            this.btnRtn.Name = "btnRtn";
            this.btnRtn.Size = new System.Drawing.Size(83, 32);
            this.btnRtn.TabIndex = 5;
            this.btnRtn.Text = "終了(&Q)";
            this.btnRtn.UseVisualStyleBackColor = true;
            this.btnRtn.Click += new System.EventHandler(this.btnRtn_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // chkTimeYukyu
            // 
            this.chkTimeYukyu.AutoSize = true;
            this.chkTimeYukyu.Location = new System.Drawing.Point(30, 615);
            this.chkTimeYukyu.Name = "chkTimeYukyu";
            this.chkTimeYukyu.Size = new System.Drawing.Size(223, 24);
            this.chkTimeYukyu.TabIndex = 12;
            this.chkTimeYukyu.Text = "スタッフ時間単位有休処理を行う";
            this.chkTimeYukyu.UseVisualStyleBackColor = true;
            // 
            // frmConfig
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(645, 651);
            this.Controls.Add(this.chkTimeYukyu);
            this.Controls.Add(this.btnRtn);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnScanner);
            this.Controls.Add(this.txtScanner);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtMonth);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtYear);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("メイリオ", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.Name = "frmConfig";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "環境設定";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmConfig_FormClosing);
            this.Load += new System.EventHandler(this.frmConfig_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtYear;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMonth;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtScanner;
        private System.Windows.Forms.Button btnScanner;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ComboBox cmbBkdels;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button btnTif;
        private System.Windows.Forms.TextBox txtTif;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnDat;
        private System.Windows.Forms.TextBox txtDat;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button btnMsts;
        private System.Windows.Forms.TextBox txtMsts;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.ComboBox cmbBkdelp;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Button BtnTif2;
        private System.Windows.Forms.TextBox txtTif2;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btnDat2;
        private System.Windows.Forms.TextBox txtDat2;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button btnMstj;
        private System.Windows.Forms.TextBox txtMstj;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnRtn;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.CheckBox chkTimeYukyu;
    }
}