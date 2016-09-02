﻿namespace FileMatch
{
    partial class frmTidy
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTidy));
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox_Grade = new System.Windows.Forms.GroupBox();
            this.radioButton2 = new System.Windows.Forms.RadioButton();
            this.radioButton3 = new System.Windows.Forms.RadioButton();
            this.radioButton4 = new System.Windows.Forms.RadioButton();
            this.radioButton1 = new System.Windows.Forms.RadioButton();
            this.groupBox_Year = new System.Windows.Forms.GroupBox();
            this.cb_year = new System.Windows.Forms.CheckBox();
            this.num_year = new System.Windows.Forms.NumericUpDown();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.cb_DelayDate = new System.Windows.Forms.CheckBox();
            this.cb_delete = new System.Windows.Forms.CheckBox();
            this.cb_secret = new System.Windows.Forms.CheckBox();
            this.cb_signature = new System.Windows.Forms.CheckBox();
            this.cb_authoration = new System.Windows.Forms.CheckBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.cb_explain = new System.Windows.Forms.ComboBox();
            this.radioButton7 = new System.Windows.Forms.RadioButton();
            this.radioButton6 = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.radioButton5 = new System.Windows.Forms.RadioButton();
            this.axCAJAX1 = new AxCAJAXLib.AxCAJAX();
            this.listView2 = new System.Windows.Forms.ListView();
            this.columnHeader7 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tb_TaskCode = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tb_School = new System.Windows.Forms.TextBox();
            this.tb_Code = new System.Windows.Forms.TextBox();
            this.HideCodeLabel = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader4 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader5 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader8 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.label5 = new System.Windows.Forms.Label();
            this.label_1 = new System.Windows.Forms.Label();
            this.label_2 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.文件名 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.整理路径 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.顺序 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.路径 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.groupBox_Grade.SuspendLayout();
            this.groupBox_Year.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.num_year)).BeginInit();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axCAJAX1)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(5, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 12);
            this.label3.TabIndex = 54;
            this.label3.Text = "当前任务数/总任务数:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 12);
            this.label2.TabIndex = 53;
            this.label2.Text = "当前页/非正文页总页数:";
            // 
            // groupBox_Grade
            // 
            this.groupBox_Grade.Controls.Add(this.radioButton2);
            this.groupBox_Grade.Controls.Add(this.radioButton3);
            this.groupBox_Grade.Controls.Add(this.radioButton4);
            this.groupBox_Grade.Controls.Add(this.radioButton1);
            this.groupBox_Grade.Location = new System.Drawing.Point(5, 277);
            this.groupBox_Grade.Name = "groupBox_Grade";
            this.groupBox_Grade.Size = new System.Drawing.Size(206, 78);
            this.groupBox_Grade.TabIndex = 50;
            this.groupBox_Grade.TabStop = false;
            this.groupBox_Grade.Text = "论文级别";
            // 
            // radioButton2
            // 
            this.radioButton2.AutoSize = true;
            this.radioButton2.Location = new System.Drawing.Point(96, 26);
            this.radioButton2.Name = "radioButton2";
            this.radioButton2.Size = new System.Drawing.Size(47, 16);
            this.radioButton2.TabIndex = 8;
            this.radioButton2.Text = "博士";
            this.radioButton2.UseVisualStyleBackColor = true;
            this.radioButton2.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // radioButton3
            // 
            this.radioButton3.AutoSize = true;
            this.radioButton3.Location = new System.Drawing.Point(25, 49);
            this.radioButton3.Name = "radioButton3";
            this.radioButton3.Size = new System.Drawing.Size(59, 16);
            this.radioButton3.TabIndex = 9;
            this.radioButton3.Text = "博士后";
            this.radioButton3.UseVisualStyleBackColor = true;
            this.radioButton3.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // radioButton4
            // 
            this.radioButton4.AutoSize = true;
            this.radioButton4.Location = new System.Drawing.Point(96, 49);
            this.radioButton4.Name = "radioButton4";
            this.radioButton4.Size = new System.Drawing.Size(47, 16);
            this.radioButton4.TabIndex = 10;
            this.radioButton4.Text = "待定";
            this.radioButton4.UseVisualStyleBackColor = true;
            this.radioButton4.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // radioButton1
            // 
            this.radioButton1.AutoSize = true;
            this.radioButton1.Checked = true;
            this.radioButton1.ForeColor = System.Drawing.Color.Red;
            this.radioButton1.Location = new System.Drawing.Point(26, 26);
            this.radioButton1.Name = "radioButton1";
            this.radioButton1.Size = new System.Drawing.Size(47, 16);
            this.radioButton1.TabIndex = 7;
            this.radioButton1.TabStop = true;
            this.radioButton1.Text = "硕士";
            this.radioButton1.UseVisualStyleBackColor = true;
            this.radioButton1.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // groupBox_Year
            // 
            this.groupBox_Year.Controls.Add(this.cb_year);
            this.groupBox_Year.Controls.Add(this.num_year);
            this.groupBox_Year.Location = new System.Drawing.Point(5, 215);
            this.groupBox_Year.Name = "groupBox_Year";
            this.groupBox_Year.Size = new System.Drawing.Size(206, 61);
            this.groupBox_Year.TabIndex = 49;
            this.groupBox_Year.TabStop = false;
            this.groupBox_Year.Text = "学位年度";
            // 
            // cb_year
            // 
            this.cb_year.AutoSize = true;
            this.cb_year.Location = new System.Drawing.Point(105, 29);
            this.cb_year.Name = "cb_year";
            this.cb_year.Size = new System.Drawing.Size(48, 16);
            this.cb_year.TabIndex = 6;
            this.cb_year.Text = "待定";
            this.cb_year.UseVisualStyleBackColor = true;
            this.cb_year.CheckedChanged += new System.EventHandler(this.cb_year_CheckedChanged_1);
            // 
            // num_year
            // 
            this.num_year.Location = new System.Drawing.Point(15, 26);
            this.num_year.Maximum = new decimal(new int[] {
            2040,
            0,
            0,
            0});
            this.num_year.Name = "num_year";
            this.num_year.Size = new System.Drawing.Size(76, 21);
            this.num_year.TabIndex = 5;
            this.num_year.Value = new decimal(new int[] {
            2015,
            0,
            0,
            0});
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.cb_DelayDate);
            this.groupBox4.Controls.Add(this.cb_delete);
            this.groupBox4.Controls.Add(this.cb_secret);
            this.groupBox4.Location = new System.Drawing.Point(5, 356);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(206, 88);
            this.groupBox4.TabIndex = 52;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "基本信息";
            // 
            // cb_DelayDate
            // 
            this.cb_DelayDate.AutoSize = true;
            this.cb_DelayDate.Location = new System.Drawing.Point(88, 53);
            this.cb_DelayDate.Name = "cb_DelayDate";
            this.cb_DelayDate.Size = new System.Drawing.Size(72, 16);
            this.cb_DelayDate.TabIndex = 5;
            this.cb_DelayDate.Text = "滞后上网";
            this.cb_DelayDate.UseVisualStyleBackColor = true;
            this.cb_DelayDate.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // cb_delete
            // 
            this.cb_delete.AutoSize = true;
            this.cb_delete.Location = new System.Drawing.Point(8, 53);
            this.cb_delete.Name = "cb_delete";
            this.cb_delete.Size = new System.Drawing.Size(72, 16);
            this.cb_delete.TabIndex = 3;
            this.cb_delete.Text = "删除字样";
            this.cb_delete.UseVisualStyleBackColor = true;
            this.cb_delete.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // cb_secret
            // 
            this.cb_secret.AutoSize = true;
            this.cb_secret.ForeColor = System.Drawing.SystemColors.ControlText;
            this.cb_secret.Location = new System.Drawing.Point(8, 28);
            this.cb_secret.Name = "cb_secret";
            this.cb_secret.Size = new System.Drawing.Size(48, 16);
            this.cb_secret.TabIndex = 1;
            this.cb_secret.Text = "保密";
            this.cb_secret.UseVisualStyleBackColor = true;
            this.cb_secret.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // cb_signature
            // 
            this.cb_signature.AutoSize = true;
            this.cb_signature.Location = new System.Drawing.Point(91, 24);
            this.cb_signature.Name = "cb_signature";
            this.cb_signature.Size = new System.Drawing.Size(84, 16);
            this.cb_signature.TabIndex = 4;
            this.cb_signature.Text = "无作者签名";
            this.cb_signature.UseVisualStyleBackColor = true;
            this.cb_signature.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // cb_authoration
            // 
            this.cb_authoration.AutoSize = true;
            this.cb_authoration.Location = new System.Drawing.Point(14, 24);
            this.cb_authoration.Name = "cb_authoration";
            this.cb_authoration.Size = new System.Drawing.Size(60, 16);
            this.cb_authoration.TabIndex = 2;
            this.cb_authoration.Text = "无授权";
            this.cb_authoration.UseVisualStyleBackColor = true;
            this.cb_authoration.CheckStateChanged += new System.EventHandler(this.cb_authoration_CheckStateChanged);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.cb_explain);
            this.groupBox5.Controls.Add(this.radioButton7);
            this.groupBox5.Controls.Add(this.cb_signature);
            this.groupBox5.Controls.Add(this.radioButton6);
            this.groupBox5.Controls.Add(this.cb_authoration);
            this.groupBox5.Controls.Add(this.label1);
            this.groupBox5.Controls.Add(this.radioButton5);
            this.groupBox5.Location = new System.Drawing.Point(5, 446);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(206, 107);
            this.groupBox5.TabIndex = 51;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "授权情况";
            // 
            // cb_explain
            // 
            this.cb_explain.Enabled = false;
            this.cb_explain.FormattingEnabled = true;
            this.cb_explain.Items.AddRange(new object[] {
            "待反馈1",
            "待反馈2",
            "待反馈3",
            "待反馈4"});
            this.cb_explain.Location = new System.Drawing.Point(124, 75);
            this.cb_explain.Name = "cb_explain";
            this.cb_explain.Size = new System.Drawing.Size(72, 20);
            this.cb_explain.TabIndex = 14;
            // 
            // radioButton7
            // 
            this.radioButton7.AutoSize = true;
            this.radioButton7.Location = new System.Drawing.Point(91, 49);
            this.radioButton7.Name = "radioButton7";
            this.radioButton7.Size = new System.Drawing.Size(59, 16);
            this.radioButton7.TabIndex = 13;
            this.radioButton7.Text = "不合格";
            this.radioButton7.UseVisualStyleBackColor = true;
            this.radioButton7.CheckedChanged += new System.EventHandler(this.radioButton7_CheckedChanged);
            // 
            // radioButton6
            // 
            this.radioButton6.AutoSize = true;
            this.radioButton6.Location = new System.Drawing.Point(15, 75);
            this.radioButton6.Name = "radioButton6";
            this.radioButton6.Size = new System.Drawing.Size(59, 16);
            this.radioButton6.TabIndex = 12;
            this.radioButton6.Text = "待反馈";
            this.radioButton6.UseVisualStyleBackColor = true;
            this.radioButton6.CheckedChanged += new System.EventHandler(this.radioButton6_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(89, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 45;
            this.label1.Text = "备注";
            // 
            // radioButton5
            // 
            this.radioButton5.AutoSize = true;
            this.radioButton5.Checked = true;
            this.radioButton5.ForeColor = System.Drawing.Color.Red;
            this.radioButton5.Location = new System.Drawing.Point(15, 49);
            this.radioButton5.Name = "radioButton5";
            this.radioButton5.Size = new System.Drawing.Size(47, 16);
            this.radioButton5.TabIndex = 11;
            this.radioButton5.TabStop = true;
            this.radioButton5.Text = "合格";
            this.radioButton5.UseVisualStyleBackColor = true;
            this.radioButton5.CheckedChanged += new System.EventHandler(this.checkOrRadioButton_CheckedChanged);
            // 
            // axCAJAX1
            // 
            this.axCAJAX1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.axCAJAX1.Enabled = true;
            this.axCAJAX1.Location = new System.Drawing.Point(187, 0);
            this.axCAJAX1.Name = "axCAJAX1";
            this.axCAJAX1.OcxState = ((System.Windows.Forms.AxHost.State)(resources.GetObject("axCAJAX1.OcxState")));
            this.axCAJAX1.Size = new System.Drawing.Size(609, 737);
            this.axCAJAX1.TabIndex = 20;
            this.axCAJAX1.MouseWheelEvent += new AxCAJAXLib._DCAJAXEvents_MouseWheelEventHandler(this.axCAJAX1_MouseWheelEvent);
            this.axCAJAX1.PageChanged += new AxCAJAXLib._DCAJAXEvents_PageChangedEventHandler(this.axCAJAX1_PageChanged);
            // 
            // listView2
            // 
            this.listView2.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader7,
            this.columnHeader2});
            this.listView2.Location = new System.Drawing.Point(3, 2);
            this.listView2.Name = "listView2";
            this.listView2.Size = new System.Drawing.Size(178, 454);
            this.listView2.SmallImageList = this.imageList1;
            this.listView2.TabIndex = 19;
            this.listView2.UseCompatibleStateImageBehavior = false;
            this.listView2.View = System.Windows.Forms.View.Details;
            this.listView2.DoubleClick += new System.EventHandler(this.listView2_DoubleClick);
            // 
            // columnHeader7
            // 
            this.columnHeader7.Text = "编号";
            this.columnHeader7.Width = 120;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "success.ico");
            this.imageList1.Images.SetKeyName(1, "fail.ico");
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tb_TaskCode);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.tb_School);
            this.groupBox1.Controls.Add(this.tb_Code);
            this.groupBox1.Controls.Add(this.HideCodeLabel);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.groupBox1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(0)))), ((int)(((byte)(64)))));
            this.groupBox1.Location = new System.Drawing.Point(5, 81);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(206, 132);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "任务信息";
            // 
            // tb_TaskCode
            // 
            this.tb_TaskCode.Enabled = false;
            this.tb_TaskCode.Location = new System.Drawing.Point(66, 67);
            this.tb_TaskCode.Name = "tb_TaskCode";
            this.tb_TaskCode.Size = new System.Drawing.Size(121, 21);
            this.tb_TaskCode.TabIndex = 37;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 70);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(53, 12);
            this.label8.TabIndex = 36;
            this.label8.Text = "接收编号";
            // 
            // tb_School
            // 
            this.tb_School.Enabled = false;
            this.tb_School.Location = new System.Drawing.Point(67, 29);
            this.tb_School.Name = "tb_School";
            this.tb_School.Size = new System.Drawing.Size(120, 21);
            this.tb_School.TabIndex = 34;
            // 
            // tb_Code
            // 
            this.tb_Code.Enabled = false;
            this.tb_Code.Location = new System.Drawing.Point(66, 102);
            this.tb_Code.Name = "tb_Code";
            this.tb_Code.Size = new System.Drawing.Size(124, 21);
            this.tb_Code.TabIndex = 25;
            // 
            // HideCodeLabel
            // 
            this.HideCodeLabel.AutoSize = true;
            this.HideCodeLabel.Location = new System.Drawing.Point(343, 421);
            this.HideCodeLabel.Name = "HideCodeLabel";
            this.HideCodeLabel.Size = new System.Drawing.Size(0, 12);
            this.HideCodeLabel.TabIndex = 23;
            this.HideCodeLabel.Visible = false;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(8, 32);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(53, 12);
            this.label9.TabIndex = 12;
            this.label9.Text = "授予单位";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(12, 106);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 12);
            this.label7.TabIndex = 10;
            this.label7.Text = "编号";
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "文件名";
            this.columnHeader1.Width = 246;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "可读否";
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "归类否";
            // 
            // columnHeader8
            // 
            this.columnHeader8.Text = "非正文页提取";
            this.columnHeader8.Width = 84;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(923, 105);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(0, 12);
            this.label5.TabIndex = 18;
            this.label5.Visible = false;
            // 
            // label_1
            // 
            this.label_1.AutoSize = true;
            this.label_1.Location = new System.Drawing.Point(148, 20);
            this.label_1.Name = "label_1";
            this.label_1.Size = new System.Drawing.Size(0, 12);
            this.label_1.TabIndex = 56;
            // 
            // label_2
            // 
            this.label_2.AutoSize = true;
            this.label_2.Location = new System.Drawing.Point(148, 49);
            this.label_2.Name = "label_2";
            this.label_2.Size = new System.Drawing.Size(0, 12);
            this.label_2.TabIndex = 57;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.label_1);
            this.groupBox2.Controls.Add(this.label_2);
            this.groupBox2.Location = new System.Drawing.Point(5, 5);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(206, 75);
            this.groupBox2.TabIndex = 61;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "实时进度";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.文件名,
            this.整理路径,
            this.顺序,
            this.路径});
            this.dataGridView1.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.dataGridView1.Location = new System.Drawing.Point(3, 458);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(183, 400);
            this.dataGridView1.TabIndex = 62;
            this.dataGridView1.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_CellMouseDoubleClick);
            this.dataGridView1.CellMouseEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellMouseEnter);
            // 
            // 文件名
            // 
            this.文件名.DataPropertyName = "文件名";
            this.文件名.HeaderText = "文件名";
            this.文件名.Name = "文件名";
            this.文件名.ReadOnly = true;
            this.文件名.Width = 125;
            // 
            // 整理路径
            // 
            this.整理路径.DataPropertyName = "整理路径";
            this.整理路径.HeaderText = "整理路径";
            this.整理路径.Name = "整理路径";
            this.整理路径.Visible = false;
            // 
            // 顺序
            // 
            this.顺序.DataPropertyName = "顺序";
            this.顺序.HeaderText = "顺序";
            this.顺序.Name = "顺序";
            this.顺序.Width = 55;
            // 
            // 路径
            // 
            this.路径.DataPropertyName = "路径";
            this.路径.HeaderText = "路径";
            this.路径.Name = "路径";
            this.路径.Visible = false;
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.groupBox4);
            this.panel1.Controls.Add(this.groupBox_Year);
            this.panel1.Controls.Add(this.groupBox_Grade);
            this.panel1.Controls.Add(this.groupBox5);
            this.panel1.Location = new System.Drawing.Point(802, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(220, 571);
            this.panel1.TabIndex = 64;
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button2.Location = new System.Drawing.Point(817, 596);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(178, 29);
            this.button2.TabIndex = 66;
            this.button2.Text = "当前任务置不可做";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(817, 699);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(178, 29);
            this.button1.TabIndex = 65;
            this.button1.Text = "设置完成标记";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button3.Location = new System.Drawing.Point(817, 647);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(178, 29);
            this.button3.TabIndex = 67;
            this.button3.Text = "任务存疑";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // frmTidy
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1020, 737);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.listView2);
            this.Controls.Add(this.axCAJAX1);
            this.Controls.Add(this.label5);
            this.Name = "frmTidy";
            this.Text = "信息标记";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmTidy_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmTidy_FormClosed);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox_Grade.ResumeLayout(false);
            this.groupBox_Grade.PerformLayout();
            this.groupBox_Year.ResumeLayout(false);
            this.groupBox_Year.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.num_year)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.axCAJAX1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        //private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tb_Code;
        private System.Windows.Forms.TextBox tb_School;
        private System.Windows.Forms.ListView listView2;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.Windows.Forms.Label HideCodeLabel;
        private System.Windows.Forms.ColumnHeader columnHeader8;
        private System.Windows.Forms.TextBox tb_TaskCode;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ImageList imageList1;
        private AxCAJAXLib.AxCAJAX axCAJAX1;
        private System.Windows.Forms.GroupBox groupBox_Grade;
        private System.Windows.Forms.RadioButton radioButton2;
        private System.Windows.Forms.RadioButton radioButton3;
        private System.Windows.Forms.RadioButton radioButton4;
        private System.Windows.Forms.RadioButton radioButton1;
        private System.Windows.Forms.GroupBox groupBox_Year;
        private System.Windows.Forms.CheckBox cb_year;
        private System.Windows.Forms.NumericUpDown num_year;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckBox cb_signature;
        private System.Windows.Forms.CheckBox cb_authoration;
        private System.Windows.Forms.CheckBox cb_secret;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.ComboBox cb_explain;
        private System.Windows.Forms.RadioButton radioButton7;
        private System.Windows.Forms.RadioButton radioButton6;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton radioButton5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label_1;
        private System.Windows.Forms.Label label_2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 文件名;
        private System.Windows.Forms.DataGridViewTextBoxColumn 整理路径;
        private System.Windows.Forms.DataGridViewTextBoxColumn 顺序;
        private System.Windows.Forms.DataGridViewTextBoxColumn 路径;
        private System.Windows.Forms.CheckBox cb_delete;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox cb_DelayDate;
        private System.Windows.Forms.Button button3;

    }
}