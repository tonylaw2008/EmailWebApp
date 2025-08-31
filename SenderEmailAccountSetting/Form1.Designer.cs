namespace SenderEmailAccountSetting
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            lblSenderOfCompany = new Label();
            txtSenderOfCompany = new TextBox();
            lblMailTool = new Label();
            chkMailTool = new CheckBox();
            label1 = new Label();
            chkEnableSSL = new CheckBox();
            label2 = new Label();
            chkEnableTSL = new CheckBox();
            label3 = new Label();
            chkEnablePasswordAuthentication = new CheckBox();
            label4 = new Label();
            txtSenderServerHost = new TextBox();
            label5 = new Label();
            txtSenderServerHostPort = new TextBox();
            label6 = new Label();
            label7 = new Label();
            label8 = new Label();
            label10 = new Label();
            label11 = new Label();
            txtRemarks = new TextBox();
            txtSenderUserPassword = new TextBox();
            txtSenderUserName = new TextBox();
            txtFromMailDisplayName = new TextBox();
            txtFromMailAddress = new TextBox();
            richTextBox1 = new RichTextBox();
            button1 = new Button();
            toolTip1 = new ToolTip(components);
            chkAddToList = new CheckBox();
            button2 = new Button();
            button3 = new Button();
            label9 = new Label();
            button4 = new Button();
            saveFileDialog1 = new SaveFileDialog();
            SuspendLayout();
            // 
            // lblSenderOfCompany
            // 
            lblSenderOfCompany.AutoSize = true;
            lblSenderOfCompany.Location = new Point(32, 39);
            lblSenderOfCompany.Name = "lblSenderOfCompany";
            lblSenderOfCompany.Size = new Size(122, 15);
            lblSenderOfCompany.TabIndex = 0;
            lblSenderOfCompany.Text = "Sender Of Company";
            // 
            // txtSenderOfCompany
            // 
            txtSenderOfCompany.AutoCompleteCustomSource.AddRange(new string[] { "gmail.com", "126.com", "yahoo.com.hk" });
            txtSenderOfCompany.Location = new Point(187, 31);
            txtSenderOfCompany.Name = "txtSenderOfCompany";
            txtSenderOfCompany.Size = new Size(157, 23);
            txtSenderOfCompany.TabIndex = 1;
            // 
            // lblMailTool
            // 
            lblMailTool.AutoSize = true;
            lblMailTool.Location = new Point(32, 67);
            lblMailTool.Name = "lblMailTool";
            lblMailTool.Size = new Size(64, 15);
            lblMailTool.TabIndex = 2;
            lblMailTool.Text = "Mail Tool ";
            // 
            // chkMailTool
            // 
            chkMailTool.AutoSize = true;
            chkMailTool.Location = new Point(187, 67);
            chkMailTool.Name = "chkMailTool";
            chkMailTool.Size = new Size(176, 19);
            chkMailTool.TabIndex = 3;
            chkMailTool.Text = "Check For MailKit ★★★★★";
            toolTip1.SetToolTip(chkMailTool, "勾選將使用MailKit發送，則 優先用SSL CER,TSL 為false，Enable Authentication 是必選的。");
            chkMailTool.UseVisualStyleBackColor = true;
            chkMailTool.CheckedChanged += chkMailTool_CheckedChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(32, 96);
            label1.Name = "label1";
            label1.Size = new Size(69, 15);
            label1.TabIndex = 4;
            label1.Text = "Enable SSL";
            // 
            // chkEnableSSL
            // 
            chkEnableSSL.AutoSize = true;
            chkEnableSSL.Location = new Point(187, 96);
            chkEnableSSL.Name = "chkEnableSSL";
            chkEnableSSL.Size = new Size(104, 19);
            chkEnableSSL.TabIndex = 5;
            chkEnableSSL.Text = "Check For SSL";
            chkEnableSSL.UseVisualStyleBackColor = true;
            chkEnableSSL.CheckedChanged += chkEnableSSL_CheckedChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(32, 125);
            label2.Name = "label2";
            label2.Size = new Size(69, 15);
            label2.TabIndex = 6;
            label2.Text = "Enable TSL";
            // 
            // chkEnableTSL
            // 
            chkEnableTSL.AutoSize = true;
            chkEnableTSL.Location = new Point(187, 125);
            chkEnableTSL.Name = "chkEnableTSL";
            chkEnableTSL.Size = new Size(104, 19);
            chkEnableTSL.TabIndex = 7;
            chkEnableTSL.Text = "Check For TSL";
            chkEnableTSL.UseVisualStyleBackColor = true;
            chkEnableTSL.CheckedChanged += chkEnableTSL_CheckedChanged;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(32, 158);
            label3.Name = "label3";
            label3.Size = new Size(131, 15);
            label3.TabIndex = 8;
            label3.Text = "Enable Authentication";
            // 
            // chkEnablePasswordAuthentication
            // 
            chkEnablePasswordAuthentication.AutoSize = true;
            chkEnablePasswordAuthentication.CausesValidation = false;
            chkEnablePasswordAuthentication.Location = new Point(187, 154);
            chkEnablePasswordAuthentication.Name = "chkEnablePasswordAuthentication";
            chkEnablePasswordAuthentication.Size = new Size(261, 19);
            chkEnablePasswordAuthentication.TabIndex = 9;
            chkEnablePasswordAuthentication.Text = "Check For  enablePasswordAuthentication";
            chkEnablePasswordAuthentication.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(32, 195);
            label4.Name = "label4";
            label4.Size = new Size(76, 15);
            label4.TabIndex = 10;
            label4.Text = "Sender Host";
            // 
            // txtSenderServerHost
            // 
            txtSenderServerHost.Location = new Point(187, 187);
            txtSenderServerHost.Name = "txtSenderServerHost";
            txtSenderServerHost.Size = new Size(157, 23);
            txtSenderServerHost.TabIndex = 11;
            toolTip1.SetToolTip(txtSenderServerHost, "like this: smtp.gmali.com");
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(32, 224);
            label5.Name = "label5";
            label5.Size = new Size(73, 15);
            label5.TabIndex = 12;
            label5.Text = "Sender Port";
            // 
            // txtSenderServerHostPort
            // 
            txtSenderServerHostPort.AutoCompleteCustomSource.AddRange(new string[] { "465", "587" });
            txtSenderServerHostPort.Location = new Point(187, 217);
            txtSenderServerHostPort.Name = "txtSenderServerHostPort";
            txtSenderServerHostPort.Size = new Size(157, 23);
            txtSenderServerHostPort.TabIndex = 13;
            toolTip1.SetToolTip(txtSenderServerHostPort, " SSL 默認PORT465 | TLS 默認PORT 587");
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Location = new Point(32, 257);
            label6.Name = "label6";
            label6.Size = new Size(106, 15);
            label6.TabIndex = 14;
            label6.Text = "FromMailAddress";
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Location = new Point(32, 290);
            label7.Name = "label7";
            label7.Size = new Size(137, 15);
            label7.TabIndex = 15;
            label7.Text = "FromMailDisplayName";
            // 
            // label8
            // 
            label8.AutoSize = true;
            label8.Location = new Point(32, 323);
            label8.Name = "label8";
            label8.Size = new Size(107, 15);
            label8.TabIndex = 16;
            label8.Text = "SenderUserName";
            // 
            // label10
            // 
            label10.AutoSize = true;
            label10.Location = new Point(32, 385);
            label10.Name = "label10";
            label10.Size = new Size(55, 15);
            label10.TabIndex = 18;
            label10.Text = "Remarks";
            // 
            // label11
            // 
            label11.AutoSize = true;
            label11.Location = new Point(32, 359);
            label11.Name = "label11";
            label11.Size = new Size(125, 15);
            label11.TabIndex = 19;
            label11.Text = "SenderUserPassword";
            // 
            // txtRemarks
            // 
            txtRemarks.AutoCompleteCustomSource.AddRange(new string[] { "[mailTool:System.Net.Mail.SmtpClient=0; MailKit.Net.Smtp.SmtpClient=1] [StartTLS=587;SSL=465] This is the brevo.com email account used for sending emails from MyCompany." });
            txtRemarks.Location = new Point(187, 385);
            txtRemarks.Name = "txtRemarks";
            txtRemarks.Size = new Size(261, 23);
            txtRemarks.TabIndex = 23;
            // 
            // txtSenderUserPassword
            // 
            txtSenderUserPassword.Location = new Point(187, 353);
            txtSenderUserPassword.Name = "txtSenderUserPassword";
            txtSenderUserPassword.Size = new Size(157, 23);
            txtSenderUserPassword.TabIndex = 24;
            toolTip1.SetToolTip(txtSenderUserPassword, "Smtp Server login Password");
            // 
            // txtSenderUserName
            // 
            txtSenderUserName.Location = new Point(187, 319);
            txtSenderUserName.Name = "txtSenderUserName";
            txtSenderUserName.Size = new Size(157, 23);
            txtSenderUserName.TabIndex = 25;
            toolTip1.SetToolTip(txtSenderUserName, "Smtp Server login Account, Like This :  ServiceCenter@MyCompany.com");
            // 
            // txtFromMailDisplayName
            // 
            txtFromMailDisplayName.AutoCompleteCustomSource.AddRange(new string[] { "My Compny Service Department" });
            txtFromMailDisplayName.Location = new Point(187, 284);
            txtFromMailDisplayName.Name = "txtFromMailDisplayName";
            txtFromMailDisplayName.Size = new Size(157, 23);
            txtFromMailDisplayName.TabIndex = 26;
            toolTip1.SetToolTip(txtFromMailDisplayName, "Like this : Manager Jack Ma");
            // 
            // txtFromMailAddress
            // 
            txtFromMailAddress.AutoCompleteCustomSource.AddRange(new string[] { "service123123@gmail.com" });
            txtFromMailAddress.Location = new Point(187, 249);
            txtFromMailAddress.Name = "txtFromMailAddress";
            txtFromMailAddress.Size = new Size(157, 23);
            txtFromMailAddress.TabIndex = 27;
            toolTip1.SetToolTip(txtFromMailAddress, "Like This : JackMa@gmail.com");
            // 
            // richTextBox1
            // 
            richTextBox1.Location = new Point(489, 31);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.Size = new Size(596, 345);
            richTextBox1.TabIndex = 28;
            richTextBox1.Text = "";
            // 
            // button1
            // 
            button1.Location = new Point(33, 445);
            button1.Name = "button1";
            button1.Size = new Size(75, 23);
            button1.TabIndex = 29;
            button1.Text = "Generate";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // chkAddToList
            // 
            chkAddToList.AutoSize = true;
            chkAddToList.Checked = true;
            chkAddToList.CheckState = CheckState.Checked;
            chkAddToList.Location = new Point(128, 448);
            chkAddToList.Name = "chkAddToList";
            chkAddToList.Size = new Size(89, 19);
            chkAddToList.TabIndex = 33;
            chkAddToList.Text = "Add To List";
            toolTip1.SetToolTip(chkAddToList, "Save to Aspsettings.josn or SenderEmailAccountList_6000018.json");
            chkAddToList.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            button2.Location = new Point(489, 400);
            button2.Name = "button2";
            button2.Size = new Size(75, 23);
            button2.TabIndex = 30;
            button2.Text = "Demo TSL";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // button3
            // 
            button3.Location = new Point(592, 400);
            button3.Name = "button3";
            button3.Size = new Size(75, 23);
            button3.TabIndex = 31;
            button3.Text = "Demo SSL";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // label9
            // 
            label9.AutoSize = true;
            label9.Location = new Point(489, 445);
            label9.Name = "label9";
            label9.Size = new Size(468, 15);
            label9.TabIndex = 32;
            label9.Text = "MailKit tool sends SSL, .NET MAIL does not support it, and its functions are limited.";
            // 
            // button4
            // 
            button4.Location = new Point(984, 385);
            button4.Name = "button4";
            button4.Size = new Size(101, 23);
            button4.TabIndex = 34;
            button4.Text = "&Save to file";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1097, 515);
            Controls.Add(button4);
            Controls.Add(chkAddToList);
            Controls.Add(label9);
            Controls.Add(button3);
            Controls.Add(button2);
            Controls.Add(button1);
            Controls.Add(richTextBox1);
            Controls.Add(txtFromMailAddress);
            Controls.Add(txtFromMailDisplayName);
            Controls.Add(txtSenderUserName);
            Controls.Add(txtSenderUserPassword);
            Controls.Add(txtRemarks);
            Controls.Add(label11);
            Controls.Add(label10);
            Controls.Add(label8);
            Controls.Add(label7);
            Controls.Add(label6);
            Controls.Add(txtSenderServerHostPort);
            Controls.Add(label5);
            Controls.Add(txtSenderServerHost);
            Controls.Add(label4);
            Controls.Add(chkEnablePasswordAuthentication);
            Controls.Add(label3);
            Controls.Add(chkEnableTSL);
            Controls.Add(label2);
            Controls.Add(chkEnableSSL);
            Controls.Add(label1);
            Controls.Add(chkMailTool);
            Controls.Add(lblMailTool);
            Controls.Add(txtSenderOfCompany);
            Controls.Add(lblSenderOfCompany);
            Name = "Form1";
            Text = "發送EMAIL賬號配置";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label lblSenderOfCompany;
        private TextBox txtSenderOfCompany;
        private Label lblMailTool;
        private CheckBox chkMailTool;
        private Label label1;
        private CheckBox chkEnableSSL;
        private Label label2;
        private CheckBox chkEnableTSL;
        private Label label3;
        private CheckBox chkEnablePasswordAuthentication;
        private Label label4;
        private TextBox txtSenderServerHost;
        private Label label5;
        private TextBox txtSenderServerHostPort;
        private Label label6;
        private Label label7;
        private Label label8;
        private Label label10;
        private Label label11;
        private TextBox txtRemarks;
        private TextBox txtSenderUserPassword;
        private TextBox txtSenderUserName;
        private TextBox txtFromMailDisplayName;
        private TextBox txtFromMailAddress;
        private RichTextBox richTextBox1;
        private Button button1;
        private ToolTip toolTip1;
        private Button button2;
        private Button button3;
        private Label label9;
        private CheckBox chkAddToList;
        private Button button4;
        private SaveFileDialog saveFileDialog1;
    }
}
