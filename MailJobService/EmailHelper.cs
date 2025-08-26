using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace MailJobService
{
    public class TokenManagement
    {
        [JsonProperty("secret")]
        public string Secret { get; set; }

        [JsonProperty("issuer")]
        public string Issuer { get; set; }

        [JsonProperty("audience")]
        public string Audience { get; set; }

        [JsonProperty("accessExpiration")]
        public int AccessExpiration { get; set; }

        [JsonProperty("refreshExpiration")]
        public int RefreshExpiration { get; set; }
    }

    //https://www.cnblogs.com/mingmingruyuedlut/archive/2011/10/14/2212255.html  //GOOGLE ：https://support.google.com/mail/answer/7126229?hl=zh-Hant

    //https://docs.microsoft.com/zh-cn/dotnet/api/system.net.mail.smtpclient.host?view=netcore-3.1

    //smtp.gmail.com

    //需要安全資料傳輸層(SSL)：是

    //需要傳輸層安全性(TLS)：是(如果可用)

    //需要驗證：是

    //安全資料傳輸層(SSL) 通訊埠：465

    //傳輸層安全性(TLS)/STARTTLS 通訊埠：587

    //var token = Configuration.GetSection("tokenManagement").Get<TokenManagement>();
    public class SendMailInfo:ISendMailInfo
    {
        public string SenderOfCompany { get; set; }
        public bool EnableSSL { get; set; } 
        public bool EnableTSL { get; set; }
        public bool EnablePasswordAuthentication { get; set; } 
        public string SenderServerHost { get; set; }
        public int SenderServerHostPort { get; set; }
        public string FromMailAddress { get; set; }
        public string SenderUserName { get; set; }
        public string SenderUserPassword { get; set; }
    }
    public interface ISendMailInfo
    {
        public string SenderOfCompany { get; set; }
        public bool EnableSSL { get; set; } 
        public bool EnableTSL { get; set; } 
        public bool EnablePasswordAuthentication { get; set; } 
        public string SenderServerHost { get; set; }
        public int SenderServerHostPort { get; set; }
        public string FromMailAddress { get; set; }
        public string SenderUserName { get; set; }
        public string SenderUserPassword { get; set; }
    }

    public class GmailInfo:IGmailInfo
    { 
        public bool EnableSSL { get; set; }
        public bool EnableTSL { get; set; } 
        public bool EnablePasswordAuthentication { get; set; } 
        public string SenderServerHost { get; set; }
        public int SenderServerHostPort { get; set; } 
        public string senderMailAddress { get; set; }
        public string SenderUserPassword { get; set; }
    } 
    public interface IGmailInfo
    {
        public bool EnableSSL { get; set; } 
        public bool EnableTSL { get; set; }
        public bool EnablePasswordAuthentication { get; set; } 
        public string SenderServerHost { get; set; }
        public int SenderServerHostPort { get; set; }
        public string senderMailAddress { get; set; }
        public string SenderUserPassword { get; set; }
    }

    public class GlobalConfigFromMainServiceAppSetting : IGlobalConfigFromMainServiceAppSetting
    {
        public double StandardWorkDaySpan { get; set; }
        public string ConsoleRootFolder { get; set; }
        public string UploadFolder { get; set; }
        public string DataFolder { get; set; }
        public string ScheduleAndShiftCalc { get; set; }
        public string WebRootFolder { get; set; }
        public bool SychronizeSystemTime { get; set; }
        public double LunchTimeSpan { get; set; }
        public DateTime MonthlyScheduleCalc { get; set; }
        public bool ThisMonthScheduleCalc { get; set; }
        public bool SendMailDefaultSetting { get; set; }
        public bool SendMailDouble { get; set; }
        public SendMailInfo SendMailInfo { get; set; }
        public GmailInfo GmailInfo { get; set; } 
    }
    public interface IGlobalConfigFromMainServiceAppSetting
    {
        public double StandardWorkDaySpan { get; set; }
        public string ConsoleRootFolder { get; set; }
        public string UploadFolder { get; set; }
        public string DataFolder { get; set; }
        public string ScheduleAndShiftCalc { get; set; }
        public string WebRootFolder { get; set; }
        public bool SychronizeSystemTime { get; set; }
        public double LunchTimeSpan { get; set; }
        public DateTime MonthlyScheduleCalc { get; set; }
        public bool ThisMonthScheduleCalc { get; set; }
        public bool SendMailDefaultSetting { get; set; }
        public bool SendMailDouble { get; set; }
        public SendMailInfo SendMailInfo { get; set; }
        public GmailInfo GmailInfo { get; set; }  
    }

    public class MailMessageMain
    {
        public string ToMail { get; set; }
        public string Subject { get; set; }
        public string EmailBody { get; set; }
    }
    public class MailTaskJobRequest
    {
        public bool Success { get; set; }
        public SendMailInfo SendMailInfo { get; set; }
        public string[] ToMailArray { get; set; }
        public string Subject { get; set; }
        public string EmailBody { get; set; }
    } 
    public class MailListApiUrl
    { 
        public string HostApiUrl { get; set; }
        public int Port { get; set; }
        public int IntervalMinutes { get; set; }
    }

    public class EmailHelper
    { 
        private SendMailInfo SendMailInfo { get; set; }
        private MailMessageMain MailMessageMain { get; set; }

        private MailMessage MailMessageDefault = new MailMessage();
        private MailMessage MailMessage163 = new MailMessage();
        private MailMessage MailMessage126 = new MailMessage();
        private MailMessage MailMessageQQ = new MailMessage();
        private MailMessage MailMessageGOOGLE = new MailMessage();

        ///<summary>
        /// 构造函数
        ///</summary>
        ///<param name="server">发件箱的邮件服务器地址</param>
        ///<param name="toMail">收件人地址（可以是多个收件人，程序中是以“;"进行区分的）</param>
        ///<param name="fromMail">发件人地址</param>
        ///<param name="subject">邮件标题</param>
        ///<param name="emailBody">邮件内容（可以以html格式进行设计）</param>
        ///<param name="username">发件箱的用户名（即@符号前面的字符串，例如：hello@163.com，用户名为：hello）</param>
        ///<param name="password">发件人邮箱密码</param>
        ///<param name="port">发送邮件所用的端口号（htmp协议默认为25）</param>
        ///<param name="sslEnable">true表示对邮件内容进行socket层加密传输，false表示不加密</param>
        ///<param name="pwdCheckEnable">true表示对发件人邮箱进行密码验证，false表示不对发件人邮箱进行密码验证</param>
        public EmailHelper(SendMailInfo sendMailInfo, string toMail,  string subject, string emailBody)
        {
            SendMailInfo = sendMailInfo;
            MailMessageMain = new MailMessageMain { ToMail = toMail, Subject = subject, EmailBody = emailBody };

            try
            {
                Console.WriteLine("Begin EmailHelper....");  
  
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        ///<summary>
        /// 添加附件
        ///</summary>
        ///<param name="attachmentsPath">附件的路径集合，以分号分隔</param>
        public void AddAttachments(string attachmentsPath)
        {
            if(!string.IsNullOrEmpty(attachmentsPath))
            {
                try
                {
                    string[] path = attachmentsPath.Split(';'); //以什么符号分隔可以自定义
                    Attachment data;
                    ContentDisposition disposition;
                    for (int i = 0; i < path.Length; i++)
                    {
                        data = new Attachment(path[i], MediaTypeNames.Application.Octet);

                        disposition = data.ContentDisposition;
                        disposition.CreationDate = File.GetCreationTime(path[i]);
                        disposition.ModificationDate = File.GetLastWriteTime(path[i]);
                        disposition.ReadDate = File.GetLastAccessTime(path[i]);

                        switch (SendMailInfo.SenderOfCompany)
                        {
                            case "163":
                                MailMessage163.Attachments.Add(data);
                                break;
                            case "QQ":
                                MailMessageQQ.Attachments.Add(data);
                                break;
                            default:
                                MailMessageDefault.Attachments.Add(data);
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    MailCommonBase.OperateDateLoger(string.Format("[AddAttachments] [PATH :{0} [MESSAGE]:{1}", attachmentsPath, ex.Message));
                }
            }  
        }

        public void SendMail126()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            Thread.Sleep(20);
            //126 
            MailMessage126.From = new System.Net.Mail.MailAddress(SendMailInfo.FromMailAddress);
            MailMessage126.To.Add(MailMessageMain.ToMail);
#if DEBUG
            MailMessage126.Bcc.Add("caihaili82@gmail.com");
#endif
            MailMessage126.Subject = MailMessageMain.Subject;
            MailMessage126.Body = MailMessageMain.EmailBody;
            MailMessage126.IsBodyHtml = true;
            MailMessage126.SubjectEncoding = System.Text.Encoding.UTF8;
            MailMessage126.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessage126.Priority = System.Net.Mail.MailPriority.High;
            MailMessage126.Sender = new MailAddress(SendMailInfo.FromMailAddress);

            try
            {
                SmtpClient client = new SmtpClient();
                client.Host = SendMailInfo.SenderServerHost; //"smtp.126.com";
                client.UseDefaultCredentials = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new System.Net.NetworkCredential(SendMailInfo.SenderUserName, SendMailInfo.SenderUserPassword);   
                client.Send(MailMessage126);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            { 
                string loggerLine = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [126 MAIL] [TO :{1} FROM : {2}] [SUCCESS {3}Milliseconds]", DateTime.Now, MailMessageMain.ToMail, MailMessage126.Sender, sw.Elapsed.Milliseconds);
                Console.WriteLine(loggerLine);
                MailCommonBase.OperateDateLoger(loggerLine);
                MailMessage126.Dispose();
            }
            sw.Stop();
        }

        public void SendMail163()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //163=================================================================================== 
            MailMessage163 = new MailMessage();
            MailMessage163.From = new System.Net.Mail.MailAddress(SendMailInfo.FromMailAddress);
            MailMessage163.To.Add(MailMessageMain.ToMail);
#if DEBUG
            MailMessage163.Bcc.Add("caihaili82@gmail.com");
#endif
            MailMessage163.Subject = MailMessageMain.Subject;
            MailMessage163.Body = MailMessageMain.EmailBody;
            MailMessage163.IsBodyHtml = true;
            MailMessage163.SubjectEncoding = System.Text.Encoding.UTF8;
            MailMessage163.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessage163.Priority = System.Net.Mail.MailPriority.High;
            MailMessage163.Sender = new MailAddress(SendMailInfo.FromMailAddress);
            try
            {
                SmtpClient client = new SmtpClient();
                client.Host = SendMailInfo.SenderServerHost; //"smtp.163.com";//使用163的SMTP服务器发送邮件  
                client.UseDefaultCredentials = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                string senderUserName = SendMailInfo.SenderUserName;   // _SendMailInfo.SenderUserName.Split('@')[0];

                if (SendMailInfo.EnablePasswordAuthentication)
                {
                    //local machine pc
                    System.Net.NetworkCredential nc = new System.Net.NetworkCredential(senderUserName, SendMailInfo.SenderUserPassword);
                    client.Credentials = nc.GetCredential(SendMailInfo.SenderServerHost, SendMailInfo.SenderServerHostPort, "NTLM");
                }
                else
                {
                    client.Credentials = new System.Net.NetworkCredential(senderUserName, SendMailInfo.SenderUserPassword);
                }

                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(MailMessage163); 
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            { 
                string loggerLine = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [163 MAIL] [TO :{1} FROM : {2}] [SUCCESS {3}Milliseconds]", DateTime.Now, MailMessageMain.ToMail, MailMessage163.Sender, sw.Elapsed.Milliseconds);
                Console.WriteLine(loggerLine);
                MailCommonBase.OperateDateLoger(loggerLine);
                MailMessageDefault.Dispose();
            }
            sw.Stop();
        }

        public void SendMailQQ()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //QQ 
            MailMessageQQ.From = new System.Net.Mail.MailAddress(SendMailInfo.FromMailAddress);
            MailMessageQQ.To.Add(MailMessageMain.ToMail);
#if DEBUG
            MailMessageQQ.Bcc.Add("caihaili82@gmail.com");
#endif
            MailMessageQQ.Subject = MailMessageMain.Subject;
            MailMessageQQ.Body = MailMessageMain.EmailBody;
            MailMessageQQ.IsBodyHtml = true;
            MailMessageQQ.SubjectEncoding = System.Text.Encoding.UTF8;
            MailMessageQQ.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessageQQ.Priority = System.Net.Mail.MailPriority.High;
            MailMessageQQ.Sender = new MailAddress(SendMailInfo.FromMailAddress);
            try
            {
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = SendMailInfo.SenderServerHost;
                client.UseDefaultCredentials = true;
                client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential(SendMailInfo.SenderUserName, SendMailInfo.SenderUserPassword); // userName = abc pass = ***   abc@mail.com

                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(MailMessageQQ);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {  
                string loggerLine = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [QQMAIL] [TO :{1} FROM : {2}] [SUCCESS {3}Milliseconds]", DateTime.Now, MailMessageMain.ToMail, MailMessageQQ.Sender, sw.Elapsed.Milliseconds);
                Console.WriteLine(loggerLine);
                MailCommonBase.OperateDateLoger(loggerLine);
                MailMessageDefault.Dispose();
            }
            sw.Stop();
        }

        public void SendMailGoogle()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            MailMessageGOOGLE = new MailMessage();
            //default
            MailMessageGOOGLE.To.Add(MailMessageMain.ToMail);
#if DEBUG
            MailMessageGOOGLE.Bcc.Add("caihaili82@gmail.com");
#endif
            MailMessageGOOGLE.From = new MailAddress(SendMailInfo.FromMailAddress);
            MailMessageGOOGLE.Subject = MailMessageMain.Subject;
            MailMessageGOOGLE.Body = MailMessageMain.EmailBody;
            MailMessageGOOGLE.IsBodyHtml = true;
            MailMessageGOOGLE.SubjectEncoding = System.Text.Encoding.UTF8;
            MailMessageGOOGLE.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessageGOOGLE.IsBodyHtml = true;
            MailMessageGOOGLE.Priority = MailPriority.High;
            MailMessageGOOGLE.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
            MailMessageGOOGLE.Sender = new MailAddress(SendMailInfo.FromMailAddress);

            try
            {
                if (MailMessageGOOGLE != null)
                {
                    SmtpClient smtpClient = new SmtpClient();
                    smtpClient.Host = SendMailInfo.SenderServerHost;
                    smtpClient.Port = SendMailInfo.SenderServerHostPort;

                    smtpClient.UseDefaultCredentials = true; //如果使用默认凭据，则为 true；否则为 false

                    smtpClient.EnableSsl = SendMailInfo.EnableSSL;
                    SecureString SecureStringOfSenderUserPassword = SecureStringConverter(SendMailInfo.SenderUserPassword);
                    if (SendMailInfo.EnablePasswordAuthentication)
                    {
                        System.Net.NetworkCredential nc = new System.Net.NetworkCredential(SendMailInfo.SenderUserName, SendMailInfo.SenderUserPassword);
                        //NTLM: Secure Password Authentication in Microsoft Outlook Express
                        smtpClient.Credentials = nc.GetCredential(smtpClient.Host, smtpClient.Port, "NTLM");
                    }
                    else
                    {
                        smtpClient.Credentials = new System.Net.NetworkCredential(SendMailInfo.SenderUserName, SecureStringOfSenderUserPassword, SendMailInfo.SenderServerHost);
                    }
                    smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    smtpClient.Send(MailMessageGOOGLE);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                MailCommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [GMAIL] [TO :{1} [EXCEPTION]:{2}]", DateTime.Now, MailMessageMain.ToMail, ex.ToString()));
            }
            finally
            {
                string loggerLine = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [GMAIL] [TO :{1} FROM : {2}] [SUCCESS {3}Milliseconds]", DateTime.Now, MailMessageMain.ToMail, MailMessageGOOGLE.Sender, sw.Elapsed.Milliseconds);
                Console.WriteLine(loggerLine);
                MailCommonBase.OperateDateLoger(loggerLine);
                MailMessageGOOGLE.Dispose();
            }
            sw.Stop();
        }

        public void SendMail()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //default 
            MailMessageDefault.To.Add(MailMessageMain.ToMail);
#if DEBUG
            MailMessageDefault.Bcc.Add("caihaili82@gmail.com");
#endif
            MailMessageDefault.From = new MailAddress(SendMailInfo.FromMailAddress);
            MailMessageDefault.Subject = MailMessageMain.Subject;
            MailMessageDefault.Body = MailMessageMain.EmailBody;
            MailMessageDefault.IsBodyHtml = true;
            MailMessageDefault.SubjectEncoding = System.Text.Encoding.UTF8;
            MailMessageDefault.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessageDefault.IsBodyHtml = true;
            MailMessageDefault.Priority = MailPriority.High;
            MailMessageDefault.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
            MailMessageDefault.Sender = new MailAddress(SendMailInfo.FromMailAddress);

            try
            {
                System.Net.Mail.SmtpClient client = new SmtpClient();
                client.Host = SendMailInfo.SenderServerHost;
                client.UseDefaultCredentials = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new  NetworkCredential(SendMailInfo.SenderUserName, SendMailInfo.SenderUserPassword);   ////发件箱的用户名（即@符号前面的字符串，例如：hello@163.com，用户名为：hello）
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(MailMessageDefault);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                MailCommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [DefaultMAIL] [TO :{1} FROM : {2}] [SUCCESS {3}Milliseconds]", DateTime.Now, MailMessageMain.ToMail, MailMessageDefault.Sender, sw.Elapsed.Milliseconds));
                MailMessageDefault.Dispose();
            }
            sw.Stop();
        }

        //sync-------------------------------------------------------------------------
        public Task SendMailsync()
        {
            Task task = Task.Run(() => {
                SendMail();
            });
            return task;
        }
        public Task SendMail126sync()
        {
            Task task = Task.Run(() => {
                SendMail126();
            });
            return task;
        }
        public Task SendMail163sync()
        {
            Task task = Task.Run(() => {
                SendMail163();
            });
            return task;
        }
        public Task SendMailQQsync()
        {
            Task task = Task.Run(() => {
                SendMailQQ(); 
            });
            return task;
        }
        public Task SendMailGooglesync()
        {
            Task task = Task.Run(() => {
                SendMailGoogle();
            });
            return task;
        }
        public void SendEmailSync()
        {  
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network; 
            smtpClient.Host = SendMailInfo.SenderServerHost;  
            smtpClient.Credentials = new System.Net.NetworkCredential(SendMailInfo.SenderUserName, SendMailInfo.SenderUserPassword);
             
            try
            {
                Task task = Task.Run(() => {
                    smtpClient.Send(MailMessageDefault);
                }); 
            }
            catch (SmtpException ex)
            {
                File.AppendAllText("c:\\mailSend.log", ex.Message + " \r\n"); 
            }
        }
        private SecureString SecureStringConverter(string pass)
        {
            SecureString ret = new SecureString();

            foreach (char chr in pass.ToCharArray())
                ret.AppendChar(chr);

            return ret;
        }
    }

    //GMAIL =====================================================================================================================================================
    //GMAIL =====================================================================================================================================================
    public class GmailHelper
    {
        private GmailInfo GmailInfo;  
        private MailMessage GMailMessage { get; set;}
        private MailMessageMain MailMessageMain { get; set; }
        public GmailHelper(GmailInfo gmailInfo, string toMail, string subject, string emailBody)
        { 
            try
            {
                GmailInfo = gmailInfo;
                MailMessageMain = new MailMessageMain
                {
                    ToMail = toMail,
                    Subject = subject,
                    EmailBody = emailBody
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        ///<summary>
        /// 添加附件
        ///</summary>
        ///<param name="attachmentsPath">附件的路径集合，以分号分隔</param>
        public void AddAttachments(string attachmentsPath)
        {
            if (!string.IsNullOrEmpty(attachmentsPath))
            {
                try
                {
                    string[] path = attachmentsPath.Split(';'); //以什么符号分隔可以自定义
                    Attachment data;
                    ContentDisposition disposition;
                    for (int i = 0; i < path.Length; i++)
                    {
                        data = new Attachment(path[i], MediaTypeNames.Application.Octet);

                        disposition = data.ContentDisposition;
                        disposition.CreationDate = File.GetCreationTime(path[i]);
                        disposition.ModificationDate = File.GetLastWriteTime(path[i]);
                        disposition.ReadDate = File.GetLastAccessTime(path[i]);

                        GMailMessage.Attachments.Add(data);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }

        public void SendMailGoogle()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            GMailMessage = new MailMessage();
            //default 
            GMailMessage.To.Add(MailMessageMain.ToMail);
#if DEBUG
            GMailMessage.Bcc.Add("caihaili82@gmail.com");
#endif 
            GMailMessage.From = new MailAddress(GmailInfo.senderMailAddress);
            GMailMessage.Subject = MailMessageMain.Subject;
            GMailMessage.Body = MailMessageMain.EmailBody;
            GMailMessage.IsBodyHtml = true;
            GMailMessage.SubjectEncoding = System.Text.Encoding.UTF8;
            GMailMessage.BodyEncoding = System.Text.Encoding.UTF8;
            GMailMessage.IsBodyHtml = true;
            GMailMessage.Priority = MailPriority.High;
            GMailMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
            GMailMessage.Sender = new MailAddress(GmailInfo.senderMailAddress);
           
            try
            {
                if (GMailMessage != null) 
                {
                    SmtpClient smtpClient = new SmtpClient();
                    smtpClient.Host = GmailInfo.SenderServerHost;
                    smtpClient.Port = GmailInfo.SenderServerHostPort;

                    smtpClient.UseDefaultCredentials = true; //如果使用默认凭据，则为 true；否则为 false

                    smtpClient.EnableSsl = GmailInfo.EnableSSL;
                    SecureString SecureStringOfSenderUserPassword = SecureStringConverter(GmailInfo.SenderUserPassword);
                    if (GmailInfo.EnablePasswordAuthentication)
                    { 
                        System.Net.NetworkCredential nc = new System.Net.NetworkCredential(GmailInfo.senderMailAddress, GmailInfo.SenderUserPassword);
                        //NTLM: Secure Password Authentication in Microsoft Outlook Express
                        smtpClient.Credentials = nc.GetCredential(smtpClient.Host, smtpClient.Port, "NTLM");
                    }
                    else
                    {
                        smtpClient.Credentials = new System.Net.NetworkCredential(GmailInfo.senderMailAddress, SecureStringOfSenderUserPassword, GmailInfo.SenderServerHost);
                    }
                     
                    smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    smtpClient.Send(GMailMessage);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                MailCommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][GmailHelper.SendMailGoogle] [DETAULT MAIL] [TO :{1} [EXCEPTION]:{2}]", DateTime.Now, MailMessageMain.ToMail, ex.ToString()));
            }
            finally
            {
                MailCommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][GmailHelper] [GMAIL SendMailGoogle()] [TO :{1} FROM : {2}] [SUCCESS {3}Seconds]", DateTime.Now, MailMessageMain.ToMail, GMailMessage.Sender, sw.Elapsed.Seconds));
                GMailMessage.Dispose();
            }
            sw.Stop();
        }
         
        public Task SendMailGooglesync()
        {
            Task task = Task.Run(() => {
                SendMailGoogle();
            });
            return task;
        }

        private SecureString SecureStringConverter(string pass)
        {
            SecureString ret = new SecureString();

            foreach (char chr in pass.ToCharArray())
                ret.AppendChar(chr);

            return ret;
        }
    }
}
 
