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
            this.comboBox_NewCompany = new System.Windows.Forms.ComboBox();
            this.comboBox_NewPad = new System.Windows.Forms.ComboBox();
            this.comboBox_NewDevice = new System.Windows.Forms.ComboBox();
            this.textBox_DBAddress = new System.Windows.Forms.TextBox();
            this.textBox_DBSubMask = new System.Windows.Forms.TextBox();
            this.textBox_DBGateway = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button_AddIPInfo = new System.Windows.Forms.Button();
            this.textBox_AddDBGateway = new System.Windows.Forms.TextBox();
            this.textBox_AddDB_SubMask = new System.Windows.Forms.TextBox();
            this.textBox_AddDBAddress = new System.Windows.Forms.TextBox();
            this.comboBox_AddDBDevice = new System.Windows.Forms.ComboBox();
            this.comboBox_AddDBPad = new System.Windows.Forms.ComboBox();
            this.comboBox_AddDBCompany = new System.Windows.Forms.ComboBox();
            this.comboBox_DBFile = new System.Windows.Forms.ComboBox();
            this.button_ChangeDBFile = new System.Windows.Forms.Button();
            this.label10 = new System.Windows.Forms.Label();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.textBox_KeypressDelay = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.textBox_KeyStateChange = new System.Windows.Forms.TextBox();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btn_StartCopyFile
            // 
            this.btn_StartCopyFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_StartCopyFile.Location = new System.Drawing.Point(688, 246);
            this.btn_StartCopyFile.Name = "btn_StartCopyFile";
            this.btn_StartCopyFile.Size = new System.Drawing.Size(195, 23);
            this.btn_StartCopyFile.TabIndex = 0;
            this.btn_StartCopyFile.Text = "Start Copying File";
            this.btn_StartCopyFile.UseVisualStyleBackColor = true;
            this.btn_StartCopyFile.Click += new System.EventHandler(this.btn_StartCopyFile_Click);
            // 
            // textBox_StartCopyDelayFile
            // 
            this.textBox_StartCopyDelayFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_StartCopyDelayFile.Location = new System.Drawing.Point(741, 220);
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
            this.label1.Location = new System.Drawing.Point(745, 204);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(90, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Delay Copying for\r\n";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(799, 223);
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
            this.tabControl1.Size = new System.Drawing.Size(897, 301);
            this.tabControl1.TabIndex = 6;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.label13);
            this.tabPage1.Controls.Add(this.label14);
            this.tabPage1.Controls.Add(this.textBox_KeyStateChange);
            this.tabPage1.Controls.Add(this.label11);
            this.tabPage1.Controls.Add(this.label12);
            this.tabPage1.Controls.Add(this.textBox_KeypressDelay);
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
            this.tabPage1.Size = new System.Drawing.Size(889, 275);
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
            this.comboBox_TargetProgramCSV.Size = new System.Drawing.Size(214, 21);
            this.comboBox_TargetProgramCSV.TabIndex = 10;
            this.comboBox_TargetProgramCSV.SelectedIndexChanged += new System.EventHandler(this.comboBox_TargetProgramCSV_SelectedIndexChanged);
            // 
            // label_Status
            // 
            this.label_Status.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label_Status.AutoSize = true;
            this.label_Status.Font = new System.Drawing.Font("Microsoft Sans Serif", 13F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_Status.Location = new System.Drawing.Point(6, 246);
            this.label_Status.Name = "label_Status";
            this.label_Status.Size = new System.Drawing.Size(58, 22);
            this.label_Status.TabIndex = 9;
            this.label_Status.Text = "status";
            // 
            // btn_Cancel1
            // 
            this.btn_Cancel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_Cancel1.Enabled = false;
            this.btn_Cancel1.Location = new System.Drawing.Point(474, 246);
            this.btn_Cancel1.Name = "btn_Cancel1";
            this.btn_Cancel1.Size = new System.Drawing.Size(208, 23);
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
            this.comboBox_FileLocation.Size = new System.Drawing.Size(560, 21);
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
            this.tabPage2.Size = new System.Drawing.Size(889, 275);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Direct Copy(WIP)";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(879, 400);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(47, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "seconds";
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(825, 381);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(90, 13);
            this.label5.TabIndex = 6;
            this.label5.Text = "Delay Copying for\r\n";
            // 
            // btn_StartCopyDirect
            // 
            this.btn_StartCopyDirect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btn_StartCopyDirect.Location = new System.Drawing.Point(821, 423);
            this.btn_StartCopyDirect.Name = "btn_StartCopyDirect";
            this.btn_StartCopyDirect.Size = new System.Drawing.Size(142, 23);
            this.btn_StartCopyDirect.TabIndex = 4;
            this.btn_StartCopyDirect.Text = "Start Copying Section";
            this.btn_StartCopyDirect.UseVisualStyleBackColor = true;
            // 
            // textBox_StartCopyDelayDirect
            // 
            this.textBox_StartCopyDelayDirect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_StartCopyDelayDirect.Location = new System.Drawing.Point(821, 397);
            this.textBox_StartCopyDelayDirect.Name = "textBox_StartCopyDelayDirect";
            this.textBox_StartCopyDelayDirect.Size = new System.Drawing.Size(52, 20);
            this.textBox_StartCopyDelayDirect.TabIndex = 5;
            this.textBox_StartCopyDelayDirect.Text = "5";
            this.textBox_StartCopyDelayDirect.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.label10);
            this.tabPage3.Controls.Add(this.comboBox_DBFile);
            this.tabPage3.Controls.Add(this.button_ChangeDBFile);
            this.tabPage3.Controls.Add(this.groupBox3);
            this.tabPage3.Controls.Add(this.groupBox2);
            this.tabPage3.Controls.Add(this.groupBox1);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(889, 275);
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
            this.textBox3.TabIndex = 6;
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
            this.textBox2.TabIndex = 5;
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
            this.label_Version.Location = new System.Drawing.Point(855, 9);
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
            this.comboBox_NetworkAdapter.TabIndex = 2;
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
            this.groupBox1.Size = new System.Drawing.Size(436, 205);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Adapter Addresses";
            // 
            // button_RefreshAdapter
            // 
            this.button_RefreshAdapter.Location = new System.Drawing.Point(348, 16);
            this.button_RefreshAdapter.Name = "button_RefreshAdapter";
            this.button_RefreshAdapter.Size = new System.Drawing.Size(75, 23);
            this.button_RefreshAdapter.TabIndex = 3;
            this.button_RefreshAdapter.Text = "Refresh";
            this.button_RefreshAdapter.UseVisualStyleBackColor = true;
            this.button_RefreshAdapter.Click += new System.EventHandler(this.button_RefreshAdapter_Click);
            // 
            // button_ApplyIPChanges
            // 
            this.button_ApplyIPChanges.Location = new System.Drawing.Point(331, 176);
            this.button_ApplyIPChanges.Name = "button_ApplyIPChanges";
            this.button_ApplyIPChanges.Size = new System.Drawing.Size(99, 23);
            this.button_ApplyIPChanges.TabIndex = 7;
            this.button_ApplyIPChanges.Text = "Apply IP Changes";
            this.button_ApplyIPChanges.UseVisualStyleBackColor = true;
            this.button_ApplyIPChanges.Click += new System.EventHandler(this.button_ApplyIPChanges_Click);
            // 
            // comboBox_NewCompany
            // 
            this.comboBox_NewCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_NewCompany.FormattingEnabled = true;
            this.comboBox_NewCompany.Location = new System.Drawing.Point(10, 18);
            this.comboBox_NewCompany.Name = "comboBox_NewCompany";
            this.comboBox_NewCompany.Size = new System.Drawing.Size(121, 21);
            this.comboBox_NewCompany.TabIndex = 8;
            this.comboBox_NewCompany.SelectedIndexChanged += new System.EventHandler(this.comboBox_NewCompany_SelectedIndexChanged);
            // 
            // comboBox_NewPad
            // 
            this.comboBox_NewPad.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_NewPad.Enabled = false;
            this.comboBox_NewPad.FormattingEnabled = true;
            this.comboBox_NewPad.Location = new System.Drawing.Point(137, 18);
            this.comboBox_NewPad.Name = "comboBox_NewPad";
            this.comboBox_NewPad.Size = new System.Drawing.Size(121, 21);
            this.comboBox_NewPad.TabIndex = 9;
            this.comboBox_NewPad.SelectedIndexChanged += new System.EventHandler(this.comboBox_NewPad_SelectedIndexChanged);
            // 
            // comboBox_NewDevice
            // 
            this.comboBox_NewDevice.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_NewDevice.Enabled = false;
            this.comboBox_NewDevice.FormattingEnabled = true;
            this.comboBox_NewDevice.Location = new System.Drawing.Point(264, 18);
            this.comboBox_NewDevice.Name = "comboBox_NewDevice";
            this.comboBox_NewDevice.Size = new System.Drawing.Size(144, 21);
            this.comboBox_NewDevice.TabIndex = 10;
            this.comboBox_NewDevice.SelectedIndexChanged += new System.EventHandler(this.comboBox_NewDevice_SelectedIndexChanged);
            // 
            // textBox_DBAddress
            // 
            this.textBox_DBAddress.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_DBAddress.Location = new System.Drawing.Point(10, 48);
            this.textBox_DBAddress.Name = "textBox_DBAddress";
            this.textBox_DBAddress.Size = new System.Drawing.Size(121, 20);
            this.textBox_DBAddress.TabIndex = 11;
            // 
            // textBox_DBSubMask
            // 
            this.textBox_DBSubMask.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_DBSubMask.Location = new System.Drawing.Point(137, 48);
            this.textBox_DBSubMask.Name = "textBox_DBSubMask";
            this.textBox_DBSubMask.Size = new System.Drawing.Size(121, 20);
            this.textBox_DBSubMask.TabIndex = 12;
            // 
            // textBox_DBGateway
            // 
            this.textBox_DBGateway.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_DBGateway.Location = new System.Drawing.Point(264, 48);
            this.textBox_DBGateway.Name = "textBox_DBGateway";
            this.textBox_DBGateway.Size = new System.Drawing.Size(121, 20);
            this.textBox_DBGateway.TabIndex = 13;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.textBox_DBGateway);
            this.groupBox2.Controls.Add(this.textBox_DBSubMask);
            this.groupBox2.Controls.Add(this.textBox_DBAddress);
            this.groupBox2.Controls.Add(this.comboBox_NewDevice);
            this.groupBox2.Controls.Add(this.comboBox_NewPad);
            this.groupBox2.Controls.Add(this.comboBox_NewCompany);
            this.groupBox2.Location = new System.Drawing.Point(445, 3);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(428, 99);
            this.groupBox2.TabIndex = 20;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Select IP From Database";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button_AddIPInfo);
            this.groupBox3.Controls.Add(this.textBox_AddDBGateway);
            this.groupBox3.Controls.Add(this.textBox_AddDB_SubMask);
            this.groupBox3.Controls.Add(this.textBox_AddDBAddress);
            this.groupBox3.Controls.Add(this.comboBox_AddDBDevice);
            this.groupBox3.Controls.Add(this.comboBox_AddDBPad);
            this.groupBox3.Controls.Add(this.comboBox_AddDBCompany);
            this.groupBox3.Location = new System.Drawing.Point(445, 109);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(428, 99);
            this.groupBox3.TabIndex = 21;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Add IP to Database";
            // 
            // button_AddIPInfo
            // 
            this.button_AddIPInfo.Location = new System.Drawing.Point(309, 70);
            this.button_AddIPInfo.Name = "button_AddIPInfo";
            this.button_AddIPInfo.Size = new System.Drawing.Size(99, 23);
            this.button_AddIPInfo.TabIndex = 20;
            this.button_AddIPInfo.Text = "Add to DB";
            this.button_AddIPInfo.UseVisualStyleBackColor = true;
            this.button_AddIPInfo.Click += new System.EventHandler(this.button_AddIPInfo_Click);
            // 
            // textBox_AddDBGateway
            // 
            this.textBox_AddDBGateway.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_AddDBGateway.Location = new System.Drawing.Point(264, 48);
            this.textBox_AddDBGateway.Name = "textBox_AddDBGateway";
            this.textBox_AddDBGateway.Size = new System.Drawing.Size(121, 20);
            this.textBox_AddDBGateway.TabIndex = 19;
            // 
            // textBox_AddDB_SubMask
            // 
            this.textBox_AddDB_SubMask.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_AddDB_SubMask.Location = new System.Drawing.Point(137, 48);
            this.textBox_AddDB_SubMask.Name = "textBox_AddDB_SubMask";
            this.textBox_AddDB_SubMask.Size = new System.Drawing.Size(121, 20);
            this.textBox_AddDB_SubMask.TabIndex = 18;
            // 
            // textBox_AddDBAddress
            // 
            this.textBox_AddDBAddress.BackColor = System.Drawing.SystemColors.Window;
            this.textBox_AddDBAddress.Location = new System.Drawing.Point(10, 48);
            this.textBox_AddDBAddress.Name = "textBox_AddDBAddress";
            this.textBox_AddDBAddress.Size = new System.Drawing.Size(121, 20);
            this.textBox_AddDBAddress.TabIndex = 17;
            // 
            // comboBox_AddDBDevice
            // 
            this.comboBox_AddDBDevice.FormattingEnabled = true;
            this.comboBox_AddDBDevice.Location = new System.Drawing.Point(264, 18);
            this.comboBox_AddDBDevice.Name = "comboBox_AddDBDevice";
            this.comboBox_AddDBDevice.Size = new System.Drawing.Size(144, 21);
            this.comboBox_AddDBDevice.TabIndex = 16;
            // 
            // comboBox_AddDBPad
            // 
            this.comboBox_AddDBPad.FormattingEnabled = true;
            this.comboBox_AddDBPad.Location = new System.Drawing.Point(137, 18);
            this.comboBox_AddDBPad.Name = "comboBox_AddDBPad";
            this.comboBox_AddDBPad.Size = new System.Drawing.Size(121, 21);
            this.comboBox_AddDBPad.TabIndex = 15;
            this.comboBox_AddDBPad.SelectedIndexChanged += new System.EventHandler(this.comboBox_AddDBPad_SelectedIndexChanged);
            // 
            // comboBox_AddDBCompany
            // 
            this.comboBox_AddDBCompany.FormattingEnabled = true;
            this.comboBox_AddDBCompany.Location = new System.Drawing.Point(10, 18);
            this.comboBox_AddDBCompany.Name = "comboBox_AddDBCompany";
            this.comboBox_AddDBCompany.Size = new System.Drawing.Size(121, 21);
            this.comboBox_AddDBCompany.TabIndex = 14;
            // 
            // comboBox_DBFile
            // 
            this.comboBox_DBFile.AllowDrop = true;
            this.comboBox_DBFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox_DBFile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple;
            this.comboBox_DBFile.FormattingEnabled = true;
            this.comboBox_DBFile.Location = new System.Drawing.Point(313, 240);
            this.comboBox_DBFile.MaxDropDownItems = 10;
            this.comboBox_DBFile.Name = "comboBox_DBFile";
            this.comboBox_DBFile.Size = new System.Drawing.Size(560, 21);
            this.comboBox_DBFile.TabIndex = 23;
            // 
            // button_ChangeDBFile
            // 
            this.button_ChangeDBFile.Location = new System.Drawing.Point(185, 240);
            this.button_ChangeDBFile.Name = "button_ChangeDBFile";
            this.button_ChangeDBFile.Size = new System.Drawing.Size(122, 23);
            this.button_ChangeDBFile.TabIndex = 22;
            this.button_ChangeDBFile.Text = "Change File";
            this.button_ChangeDBFile.UseVisualStyleBackColor = true;
            this.button_ChangeDBFile.Click += new System.EventHandler(this.button_ChangeDBFile_Click);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(310, 224);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(162, 13);
            this.label10.TabIndex = 16;
            this.label10.Text = "Select Database File (Excel file): ";
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "*.csv";
            this.openFileDialog2.Filter = "Excel Sheet Files | *.csv";
            // 
            // label11
            // 
            this.label11.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(532, 223);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(64, 13);
            this.label11.TabIndex = 14;
            this.label11.Text = "Milliseconds";
            // 
            // label12
            // 
            this.label12.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(471, 204);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(84, 13);
            this.label12.TabIndex = 13;
            this.label12.Text = "Key Press Delay";
            // 
            // textBox_KeypressDelay
            // 
            this.textBox_KeypressDelay.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_KeypressDelay.Location = new System.Drawing.Point(474, 220);
            this.textBox_KeypressDelay.Name = "textBox_KeypressDelay";
            this.textBox_KeypressDelay.Size = new System.Drawing.Size(52, 20);
            this.textBox_KeypressDelay.TabIndex = 12;
            this.textBox_KeypressDelay.Text = "500";
            this.textBox_KeypressDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.textBox_KeypressDelay.TextChanged += new System.EventHandler(this.textBox_KeypressDelay_TextChanged);
            // 
            // label13
            // 
            this.label13.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(661, 223);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(64, 13);
            this.label13.TabIndex = 17;
            this.label13.Text = "Milliseconds";
            // 
            // label14
            // 
            this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(600, 204);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(123, 13);
            this.label14.TabIndex = 16;
            this.label14.Text = "Key State Change Delay";
            // 
            // textBox_KeyStateChange
            // 
            this.textBox_KeyStateChange.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox_KeyStateChange.Location = new System.Drawing.Point(603, 220);
            this.textBox_KeyStateChange.Name = "textBox_KeyStateChange";
            this.textBox_KeyStateChange.Size = new System.Drawing.Size(52, 20);
            this.textBox_KeyStateChange.TabIndex = 15;
            this.textBox_KeyStateChange.Text = "5";
            this.textBox_KeyStateChange.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(921, 325);
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
            this.tabPage3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
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
        private System.Windows.Forms.ComboBox comboBox_NewDevice;
        private System.Windows.Forms.ComboBox comboBox_NewPad;
        private System.Windows.Forms.ComboBox comboBox_NewCompany;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button button_AddIPInfo;
        private System.Windows.Forms.TextBox textBox_AddDBGateway;
        private System.Windows.Forms.TextBox textBox_AddDB_SubMask;
        private System.Windows.Forms.TextBox textBox_AddDBAddress;
        private System.Windows.Forms.ComboBox comboBox_AddDBDevice;
        private System.Windows.Forms.ComboBox comboBox_AddDBPad;
        private System.Windows.Forms.ComboBox comboBox_AddDBCompany;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox textBox_DBGateway;
        private System.Windows.Forms.TextBox textBox_DBSubMask;
        private System.Windows.Forms.TextBox textBox_DBAddress;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox comboBox_DBFile;
        private System.Windows.Forms.Button button_ChangeDBFile;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox textBox_KeypressDelay;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox textBox_KeyStateChange;
    }
}

