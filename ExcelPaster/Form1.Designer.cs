namespace ExcelPaster
{
    partial class MainForm
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
            this.btn_StartCopyFile = new System.Windows.Forms.Button();
            this.textBox_StartCopyDelayFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btn_SelectFile = new System.Windows.Forms.Button();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.label6 = new System.Windows.Forms.Label();
            this.comboBox_TargetProgramCSV = new System.Windows.Forms.ComboBox();
            this.label_Status = new System.Windows.Forms.Label();
            this.btn_Cancel1 = new System.Windows.Forms.Button();
            this.comboBox_FileLocation = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btn_StartCopyDirect = new System.Windows.Forms.Button();
            this.textBox_StartCopyDelayDirect = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.label7 = new System.Windows.Forms.Label();
            this.DefGate_Status = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.SubMask_Status = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.IPAdress_Status = new System.Windows.Forms.Label();
            this.textBox_IPAdress = new System.Windows.Forms.TextBox();
            this.label_DefaultGateway = new System.Windows.Forms.Label();
            this.label_SubnetMask = new System.Windows.Forms.Label();
            this.label_IPAddress = new System.Windows.Forms.Label();
            this.BgWorker = new System.ComponentModel.BackgroundWorker();
            this.label_Version = new System.Windows.Forms.Label();
            this.comboBox_NetworkAdapter = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button_RefreshAdapter = new System.Windows.Forms.Button();
            this.button_ApplyIPChanges = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_StartCopyFile
            // 
            this.btn_StartCopyFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_StartCopyFile.Location = new System.Drawing.Point(511, 183);
            this.btn_StartCopyFile.Name = "btn_StartCopyFile";
            this.btn_StartCopyFile.Size = new System.Drawing.Size(142, 23);
            this.btn_StartCopyFile.TabIndex = 0;
            this.btn_StartCopyFile.Text = "Start Copying File";
            this.btn_StartCopyFile.UseVisualStyleBackColor = true;
            this.btn_StartCopyFile.Click += new System.EventHandler(this.btn_StartCopyFile_Click);
            // 
            // textBox_StartCopyDelayFile
            // 
            this.textBox_StartCopyDelayFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_StartCopyDelayFile.Location = new System.Drawing.Point(511, 157);
            this.textBox_StartCopyDelayFile.Name = "textBox_StartCopyDelayFile";
            this.textBox_StartCopyDelayFile.Size = new System.Drawing.Size(52, 20);
            this.textBox_StartCopyDelayFile.TabIndex = 1;
            this.textBox_StartCopyDelayFile.Text = "5";
            this.textBox_StartCopyDelayFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_StartCopyDelayFile.TextChanged += new System.EventHandler(this.textBox_StartCopyDelayFile_TextChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(515, 141);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(90, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Delay Copying for\r\n";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(569, 160);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "seconds";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "*.csv";
            this.openFileDialog1.Filter = "Excel Sheet Files | *.csv";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // btn_SelectFile
            // 
            this.btn_SelectFile.Location = new System.Drawing.Point(28, 35);
            this.btn_SelectFile.Name = "btn_SelectFile";
            this.btn_SelectFile.Size = new System.Drawing.Size(122, 23);
            this.btn_SelectFile.TabIndex = 5;
            this.btn_SelectFile.Text = "Change File";
            this.btn_SelectFile.UseVisualStyleBackColor = true;
            this.btn_SelectFile.Click += new System.EventHandler(this.btn_SelectFile_Click);
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(12, 12);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(682, 243);
            this.tabControl1.TabIndex = 6;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.comboBox_TargetProgramCSV);
            this.tabPage1.Controls.Add(this.label_Status);
            this.tabPage1.Controls.Add(this.btn_Cancel1);
            this.tabPage1.Controls.Add(this.comboBox_FileLocation);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.btn_SelectFile);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btn_StartCopyFile);
            this.tabPage1.Controls.Add(this.textBox_StartCopyDelayFile);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(674, 217);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "CSV Copy";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(316, 11);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(210, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Select target program (Copying into what?):";
            // 
            // comboBox_TargetProgramCSV
            // 
            this.comboBox_TargetProgramCSV.AllowDrop = true;
            this.comboBox_TargetProgramCSV.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox_TargetProgramCSV.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_TargetProgramCSV.FormattingEnabled = true;
            this.comboBox_TargetProgramCSV.Items.AddRange(new object[] {
            "Notepad/Text Editor",
            "Excel",
            "PCCU"});
            this.comboBox_TargetProgramCSV.Location = new System.Drawing.Point(532, 8);
            this.comboBox_TargetProgramCSV.MaxDropDownItems = 10;
            this.comboBox_TargetProgramCSV.Name = "comboBox_TargetProgramCSV";
            this.comboBox_TargetProgramCSV.Size = new System.Drawing.Size(121, 21);
            this.comboBox_TargetProgramCSV.TabIndex = 10;
            this.comboBox_TargetProgramCSV.SelectedIndexChanged += new System.EventHandler(this.comboBox_TargetProgramCSV_SelectedIndexChanged);
            // 
            // label_Status
            // 
            this.label_Status.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label_Status.AutoSize = true;
            this.label_Status.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Status.Location = new System.Drawing.Point(6, 141);
            this.label_Status.Name = "label_Status";
            this.label_Status.Size = new System.Drawing.Size(58, 22);
            this.label_Status.TabIndex = 9;
            this.label_Status.Text = "status";
            // 
            // btn_Cancel1
            // 
            this.btn_Cancel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_Cancel1.Enabled = false;
            this.btn_Cancel1.Location = new System.Drawing.Point(363, 183);
            this.btn_Cancel1.Name = "btn_Cancel1";
            this.btn_Cancel1.Size = new System.Drawing.Size(142, 23);
            this.btn_Cancel1.TabIndex = 8;
            this.btn_Cancel1.Text = "Cancel ";
            this.btn_Cancel1.UseVisualStyleBackColor = true;
            this.btn_Cancel1.Click += new System.EventHandler(this.btn_Cancel1_Click);
            // 
            // comboBox_FileLocation
            // 
            this.comboBox_FileLocation.AllowDrop = true;
            this.comboBox_FileLocation.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox_FileLocation.FormattingEnabled = true;
            this.comboBox_FileLocation.Location = new System.Drawing.Point(186, 35);
            this.comboBox_FileLocation.MaxDropDownItems = 10;
            this.comboBox_FileLocation.Name = "comboBox_FileLocation";
            this.comboBox_FileLocation.Size = new System.Drawing.Size(467, 21);
            this.comboBox_FileLocation.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(25, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(233, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Select a .CSV File to copy into another program:";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.btn_StartCopyDirect);
            this.tabPage2.Controls.Add(this.textBox_StartCopyDelayDirect);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(674, 217);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Direct Copy(WIP)";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(664, 342);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "seconds";
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(610, 323);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(90, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "Delay Copying for\r\n";
            // 
            // btn_StartCopyDirect
            // 
            this.btn_StartCopyDirect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_StartCopyDirect.Location = new System.Drawing.Point(606, 365);
            this.btn_StartCopyDirect.Name = "btn_StartCopyDirect";
            this.btn_StartCopyDirect.Size = new System.Drawing.Size(142, 23);
            this.btn_StartCopyDirect.TabIndex = 4;
            this.btn_StartCopyDirect.Text = "Start Copying Section";
            this.btn_StartCopyDirect.UseVisualStyleBackColor = true;
            // 
            // textBox_StartCopyDelayDirect
            // 
            this.textBox_StartCopyDelayDirect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_StartCopyDelayDirect.Location = new System.Drawing.Point(606, 339);
            this.textBox_StartCopyDelayDirect.Name = "textBox_StartCopyDelayDirect";
            this.textBox_StartCopyDelayDirect.Size = new System.Drawing.Size(52, 20);
            this.textBox_StartCopyDelayDirect.TabIndex = 5;
            this.textBox_StartCopyDelayDirect.Text = "5";
            this.textBox_StartCopyDelayDirect.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.groupBox1);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(674, 217);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "IP Change";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(219, 51);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(44, 13);
            this.label7.TabIndex = 10;
            this.label7.Text = "Current:";
            // 
            // DefGate_Status
            // 
            this.DefGate_Status.AutoSize = true;
            this.DefGate_Status.Location = new System.Drawing.Point(219, 130);
            this.DefGate_Status.Name = "DefGate_Status";
            this.DefGate_Status.Size = new System.Drawing.Size(103, 13);
            this.DefGate_Status.TabIndex = 9;
            this.DefGate_Status.Text = "Current DefGateway";
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.SystemColors.Window;
            this.textBox3.Location = new System.Drawing.Point(103, 127);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(99, 20);
            this.textBox3.TabIndex = 8;
            this.textBox3.TextChanged += new System.EventHandler(this.textBox3_TextChanged);
            // 
            // SubMask_Status
            // 
            this.SubMask_Status.AutoSize = true;
            this.SubMask_Status.Location = new System.Drawing.Point(219, 105);
            this.SubMask_Status.Name = "SubMask_Status";
            this.SubMask_Status.Size = new System.Drawing.Size(104, 13);
            this.SubMask_Status.TabIndex = 7;
            this.SubMask_Status.Text = "Current SubnetMask";
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Window;
            this.textBox2.Location = new System.Drawing.Point(103, 103);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(99, 20);
            this.textBox2.TabIndex = 6;
            this.textBox2.TextChanged += new System.EventHandler(this.textBox2_TextChanged);
            // 
            // IPAdress_Status
            // 
            this.IPAdress_Status.AutoSize = true;
            this.IPAdress_Status.Location = new System.Drawing.Point(219, 79);
            this.IPAdress_Status.Name = "IPAdress_Status";
            this.IPAdress_Status.Size = new System.Drawing.Size(54, 13);
            this.IPAdress_Status.TabIndex = 5;
            this.IPAdress_Status.Text = "Current IP";
            // 
            // textBox_IPAdress
            // 
            this.textBox_IPAdress.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_IPAdress.Location = new System.Drawing.Point(103, 76);
            this.textBox_IPAdress.Name = "textBox_IPAdress";
            this.textBox_IPAdress.Size = new System.Drawing.Size(99, 20);
            this.textBox_IPAdress.TabIndex = 4;
            this.textBox_IPAdress.TextChanged += new System.EventHandler(this.textBox_IPAdress_TextChanged);
            // 
            // label_DefaultGateway
            // 
            this.label_DefaultGateway.AutoSize = true;
            this.label_DefaultGateway.Location = new System.Drawing.Point(8, 130);
            this.label_DefaultGateway.Name = "label_DefaultGateway";
            this.label_DefaultGateway.Size = new System.Drawing.Size(89, 13);
            this.label_DefaultGateway.TabIndex = 3;
            this.label_DefaultGateway.Text = "Default Gateway:";
            // 
            // label_SubnetMask
            // 
            this.label_SubnetMask.AutoSize = true;
            this.label_SubnetMask.Location = new System.Drawing.Point(8, 106);
            this.label_SubnetMask.Name = "label_SubnetMask";
            this.label_SubnetMask.Size = new System.Drawing.Size(73, 13);
            this.label_SubnetMask.TabIndex = 2;
            this.label_SubnetMask.Text = "Subnet Mask:";
            // 
            // label_IPAddress
            // 
            this.label_IPAddress.AutoSize = true;
            this.label_IPAddress.Location = new System.Drawing.Point(8, 83);
            this.label_IPAddress.Name = "label_IPAddress";
            this.label_IPAddress.Size = new System.Drawing.Size(61, 13);
            this.label_IPAddress.TabIndex = 1;
            this.label_IPAddress.Text = "IP Address:";
            // 
            // BgWorker
            // 
            this.BgWorker.WorkerReportsProgress = true;
            this.BgWorker.WorkerSupportsCancellation = true;
            this.BgWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.BgWorker_DoWork);
            this.BgWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.BgWorker_ProgressChanged);
            this.BgWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.BgWorker_Completed);
            // 
            // label_Version
            // 
            this.label_Version.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label_Version.AutoSize = true;
            this.label_Version.Location = new System.Drawing.Point(640, 9);
            this.label_Version.Name = "label_Version";
            this.label_Version.Size = new System.Drawing.Size(29, 13);
            this.label_Version.TabIndex = 10;
            this.label_Version.Text = "V1.0";
            // 
            // comboBox_NetworkAdapter
            // 
            this.comboBox_NetworkAdapter.BackColor = System.Drawing.SystemColors.Window;
            this.comboBox_NetworkAdapter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_NetworkAdapter.ForeColor = System.Drawing.SystemColors.WindowText;
            this.comboBox_NetworkAdapter.FormattingEnabled = true;
            this.comboBox_NetworkAdapter.Location = new System.Drawing.Point(144, 18);
            this.comboBox_NetworkAdapter.Name = "comboBox_NetworkAdapter";
            this.comboBox_NetworkAdapter.Size = new System.Drawing.Size(196, 21);
            this.comboBox_NetworkAdapter.TabIndex = 11;
            this.comboBox_NetworkAdapter.SelectedIndexChanged += new System.EventHandler(this.comboBox_NetworkAdapter_SelectedIndexChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(100, 51);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(45, 13);
            this.label8.TabIndex = 12;
            this.label8.Text = "New IP:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(12, 21);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(126, 13);
            this.label9.TabIndex = 13;
            this.label9.Text = "Select Network Adapter: ";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button_ApplyIPChanges);
            this.groupBox1.Controls.Add(this.button_RefreshAdapter);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.textBox_IPAdress);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label_IPAddress);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label_SubnetMask);
            this.groupBox1.Controls.Add(this.comboBox_NetworkAdapter);
            this.groupBox1.Controls.Add(this.label_DefaultGateway);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.IPAdress_Status);
            this.groupBox1.Controls.Add(this.DefGate_Status);
            this.groupBox1.Controls.Add(this.SubMask_Status);
            this.groupBox1.Controls.Add(this.textBox3);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(429, 194);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Adapter Addresses";
            // 
            // button_RefreshAdapter
            // 
            this.button_RefreshAdapter.Location = new System.Drawing.Point(348, 16);
            this.button_RefreshAdapter.Name = "button_RefreshAdapter";
            this.button_RefreshAdapter.Size = new System.Drawing.Size(75, 23);
            this.button_RefreshAdapter.TabIndex = 14;
            this.button_RefreshAdapter.Text = "Refresh";
            this.button_RefreshAdapter.UseVisualStyleBackColor = true;
            this.button_RefreshAdapter.Click += new System.EventHandler(this.button_RefreshAdapter_Click);
            // 
            // button_ApplyIPChanges
            // 
            this.button_ApplyIPChanges.Location = new System.Drawing.Point(324, 156);
            this.button_ApplyIPChanges.Name = "button_ApplyIPChanges";
            this.button_ApplyIPChanges.Size = new System.Drawing.Size(99, 23);
            this.button_ApplyIPChanges.TabIndex = 15;
            this.button_ApplyIPChanges.Text = "Apply IP Changes";
            this.button_ApplyIPChanges.UseVisualStyleBackColor = true;
            this.button_ApplyIPChanges.Click += new System.EventHandler(this.button_ApplyIPChanges_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(706, 267);
            this.Controls.Add(this.label_Version);
            this.Controls.Add(this.tabControl1);
            this.Name = "MainForm";
            this.Text = "Excel Paster";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_StartCopyFile;
        private System.Windows.Forms.TextBox textBox_StartCopyDelayFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btn_SelectFile;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.ComboBox comboBox_FileLocation;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btn_Cancel1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btn_StartCopyDirect;
        private System.Windows.Forms.TextBox textBox_StartCopyDelayDirect;
        private System.ComponentModel.BackgroundWorker BgWorker;
        private System.Windows.Forms.Label label_Status;
        private System.Windows.Forms.Label label_Version;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboBox_TargetProgramCSV;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Label label_DefaultGateway;
        private System.Windows.Forms.Label label_SubnetMask;
        private System.Windows.Forms.Label label_IPAddress;
        private System.Windows.Forms.Label DefGate_Status;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label SubMask_Status;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label IPAdress_Status;
        private System.Windows.Forms.TextBox textBox_IPAdress;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBox_NetworkAdapter;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button_RefreshAdapter;
        private System.Windows.Forms.Button button_ApplyIPChanges;
    }
}

