namespace SenderEmailAccountSetting
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void chkMailTool_CheckedChanged(object sender, EventArgs e)
        {
            //�x���� MailKit���߰l��
            if (chkMailTool.Checked)
            {

                // Enable the MailKit tool settings
                chkEnableSSL.Checked = true;
                chkEnableTSL.Checked = false;
                chkEnablePasswordAuthentication.Checked = true;
                txtSenderServerHostPort.Text = "465"; // SSL Ĭ�JPORT465 | TLS Ĭ�JPORT 587
            }
            else
            {
                // Disable the mail tool settings
                chkEnableSSL.Checked = false;
                chkEnableTSL.Checked = true;
                chkEnablePasswordAuthentication.Checked = true; 
                txtSenderServerHostPort.Text = "587";
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string demo_TSL = @"{
                          ""senderOfCompany"": ""demosetting.com"",
                          ""mailTool"": 0,          // 0: .Net.Mail, 1: MailKit������ʹ�ã�
                          ""enableSSL"": false,     // SSL�C���l�ͣ�senderServerHostPort=465
                          ""enableTSL"": true,      // TSL �l�ͣ�senderServerHostPort=587
                          ""enablePasswordAuthentication"": true,                    // �Ƿ����ܴa�J�C 
                          ""senderServerHost"": ""smtp-relay.brevo.com"",            // SMTP �ŷ������C��ַ   
                          ""senderServerHostPort"": 587, //SSL Ĭ�JPORT465 | TLS Ĭ�JPORT 587
                          ""fromMailAddress"": ""service@demosetting.com"",          // �l�����]���ַ   
                          ""fromMailDisplayName"": ""Service Center Of My Company"", // �l�����]���@ʾ���Q
                          ""senderUserName"": ""955XXX001@smtp-brevo.com"",          // SMTP �J�C����~̖
                          ""senderUserPassword"": ""demosettingPSW"",                // SMTP �J�C����ܴa
                          ""Remarks"": ""[mailTool:System.Net.Mail.SmtpClient=0; MailKit.Net.Smtp.SmtpClient=1] [StartTLS=587;SSL=465] This is the brevo.com email account used for sending emails from Company ABC.""
            }";

            richTextBox1.Text = demo_TSL;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SenderEmailAccount senderEmailAccount = new SenderEmailAccount
            {
                SenderOfCompany = txtSenderOfCompany.Text.Trim(),
                MailTool = chkMailTool.Checked ? 1 : 0, // 0: .Net.Mail, 1: MailKit������ʹ�ã�
                EnableSSL = chkEnableSSL.Checked,
                EnableTSL = chkEnableTSL.Checked,
                EnablePasswordAuthentication = chkEnablePasswordAuthentication.Checked,
                SenderServerHost = txtSenderServerHost.Text.Trim(),
                SenderServerHostPort = int.Parse(txtSenderServerHostPort.Text.Trim()),
                FromMailAddress = txtFromMailAddress.Text.Trim(),
                FromMailDisplayName = txtFromMailDisplayName.Text.Trim(),
                SenderUserName = txtSenderUserName.Text.Trim(),
                SenderUserPassword = txtSenderUserPassword.Text.Trim(),
                Remarks = txtRemarks.Text.Trim()
            };

            string json = Newtonsoft.Json.JsonConvert.SerializeObject(senderEmailAccount, Newtonsoft.Json.Formatting.Indented);
            richTextBox1.Text = json;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string demo_SSL = @"{
                          ""senderOfCompany"": ""demosetting.com"",
                          ""mailTool"": 1,       // 0: .Net.Mail, 1: MailKit������ʹ�ã�
                          ""enableSSL"": true,   // SSL�C���l�ͣ�senderServerHostPort=465
                          ""enableTSL"": false,  // TSL �l�ͣ�senderServerHostPort=587
                          ""enablePasswordAuthentication"": true,                     // �Ƿ����ܴa�J�C 
                          ""senderServerHost"": ""smtp-relay.brevo.com"",             // SMTP �ŷ������C��ַ   
                          ""senderServerHostPort"": 465, //SSL Ĭ�JPORT465 | TLS Ĭ�JPORT 587
                          ""fromMailAddress"": ""service@demosetting.com"",           // �l�����]���ַ   
                          ""fromMailDisplayName"": ""Service Center Of My Company"",  // �l�����]���@ʾ���Q
                          ""senderUserName"": ""955XXX001@smtp-brevo.com"",           // SMTP �J�C����~̖
                          ""senderUserPassword"": ""demosettingPSW"",                 // SMTP �J�C����ܴa
                          ""Remarks"": ""[mailTool:System.Net.Mail.SmtpClient=0; MailKit.Net.Smtp.SmtpClient=1] [StartTLS=587;SSL=465] This is the brevo.com email account used for sending emails from Company ABC.""
            }";

            richTextBox1.Text = demo_SSL;
        }

        private void chkEnableTSL_CheckedChanged(object sender, EventArgs e)
        {
            //�x�� TSL
            if (chkEnableTSL.Checked)
            {
                chkEnableSSL.Checked = false;
                chkEnablePasswordAuthentication.Checked = true;
                txtSenderServerHostPort.Text = "587"; // SSL Ĭ�JPORT465 | TLS Ĭ�JPORT 587
            }
            else
            {
                chkEnableSSL.Checked = true;
                chkEnablePasswordAuthentication.Checked = true;
                txtSenderServerHostPort.Text = "465";
            }
        }

        private void chkEnableSSL_CheckedChanged(object sender, EventArgs e)
        {
            //�x�� SSL
            if (chkEnableSSL.Checked)
            {
                chkEnableTSL.Checked = false;
                chkEnablePasswordAuthentication.Checked = true;
                txtSenderServerHostPort.Text = "465"; // SSL Ĭ�JPORT465 | TLS Ĭ�JPORT 587
            }
            else
            {
                chkEnableTSL.Checked = true;
                chkEnablePasswordAuthentication.Checked = true;
                txtSenderServerHostPort.Text = "587";
            }
        }
    }
}
