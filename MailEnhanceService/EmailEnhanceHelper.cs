using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace MailEnhanceService
{
    /// <summary>
    /// 允許簡單的標準化、普通化的模版： MailTemplateEnum 這個常量定義的模版可以直接使用模版，其他需要傳入郵件內容前，先生成內容。
    /// </summary>
    public enum MailTemplateEnum
    {
        REGISTER = 0,
        FORGET_PASSWORD = 1,
        NO_TEMPLATE = 2
    }
    
    /// <summary>
    /// SenderEmailAccount 類別用於儲存發送郵件的帳戶資訊的清單單元
    /// </summary>
    public class SenderEmailAccount : ISenderEmailAccount
    {
        public string SenderOfCompany { get; set; }
        public bool EnableSSL { get; set; }
        public bool EnableTSL { get; set; }
        public bool EnablePasswordAuthentication { get; set; }
        public string SenderServerHost { get; set; }
        public int SenderServerHostPort { get; set; }
        public string SenderUserName { get; set; }
        public string FromMailAddress { get; set; }
       
        public string SenderUserPassword { get; set; }
        public string Remarks { get; set; }
    }
    public interface ISenderEmailAccount
    {
        public string SenderOfCompany { get; set; }
        public bool EnableSSL { get; set; }
        public bool EnableTSL { get; set; }
        public bool EnablePasswordAuthentication { get; set; }
        public string SenderServerHost { get; set; }
        public int SenderServerHostPort { get; set; }
        public string SenderUserName { get; set; }
        public string FromMailAddress { get; set; }
        public string SenderUserPassword { get; set; }
        public string Remarks { get; set; }
    }

    public class MailMessageMain
    {
        public string ToMail { get; set; }
        public string Subject { get; set; }
        public string EmailBody { get; set; }
    }
    
    public class EmailEnhanceHelper
    {  
        private readonly ILogger<EmailEnhanceHelper> _logger;
        private readonly IList<SenderEmailAccount> _senderAccounts;

        //默認的發件箱賬戶信息為列表第一個賬戶信息 初始化，避免null導致系統錯誤提示
        private SenderEmailAccount SenderEmailAccount = new SenderEmailAccount(); 
        private MailMessageMain MailMessageMain { get; set; } = new MailMessageMain();

        private MailMessage MailMessageDefault = new MailMessage();
        private MailMessage MailMessage163 = new MailMessage();
        private MailMessage MailMessage126 = new MailMessage();
        private MailMessage MailMessageQQ = new MailMessage();
        private MailMessage MailMessageGmail = new MailMessage();
         
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
        public EmailEnhanceHelper(IList<SenderEmailAccount> senderEmailAccountList, ILogger<EmailEnhanceHelper> logger)
        {
            _senderAccounts = senderEmailAccountList;
            SenderEmailAccount = _senderAccounts[0]; //默認的發件箱賬戶信息為列表第一個賬戶信息
            _logger = logger;
            _logger.LogInformation("EmailHelper initialized"); 
        }

        ///<summary>
        /// 添加附件 需要進一步優化改寫 2025-8-13
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
                        string senderOfCompany = SenderEmailAccount.SenderOfCompany.ToLower();
                        senderOfCompany = senderOfCompany=="google" ? "gmail" : senderOfCompany; //google 轉換為 gmail
                        switch (senderOfCompany)
                        {
                            case "163":
                                MailMessage163.Attachments.Add(data);
                                break;
                            case "126":
                                MailMessage126.Attachments.Add(data);
                                break;
                            case "qq":
                                MailMessageQQ.Attachments.Add(data);
                                break;
                            case "gmail":
                                MailMessageGmail.Attachments.Add(data);
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
         

        private bool SendMail126()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            bool sendSuccess = false;
            Exception lastException = null;

            try
            {
                // 初始化郵件訊息
                MailMessage126.From = new MailAddress(SenderEmailAccount.FromMailAddress);
                MailMessage126.To.Add(MailMessageMain.ToMail);
#if DEBUG
                MailMessage126.Bcc.Add("caihaili82@gmail.com");
#endif
                MailMessage126.Subject = MailMessageMain.Subject;
                MailMessage126.Body = MailMessageMain.EmailBody;
                MailMessage126.IsBodyHtml = true;
                MailMessage126.SubjectEncoding = Encoding.UTF8;
                MailMessage126.BodyEncoding = Encoding.UTF8;
                MailMessage126.Priority = MailPriority.High;
                MailMessage126.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
                MailMessage126.Sender = new MailAddress(SenderEmailAccount.FromMailAddress);

                using (SmtpClient client = new SmtpClient())
                {
                    client.Host = SenderEmailAccount.SenderServerHost;
                    client.Port = SenderEmailAccount.SenderServerHostPort; 
                    client.EnableSsl = SenderEmailAccount.EnableSSL;
                    client.Timeout = 10000; // 設置超時10秒
                    client.UseDefaultCredentials = false;

                    // 身份驗證
                    client.Credentials = new NetworkCredential(
                        SenderEmailAccount.SenderUserName,
                        SenderEmailAccount.SenderUserPassword
                    );
                    
                    // 發送郵件
                    client.Send(MailMessage126);  //如果發送郵件持有證書令牌 則使用  client.SendAsync(MailMessage126,userToken); 
                    sendSuccess = true;

                   
                }   
            }
            catch (SmtpException smtpEx)
            {
                // 發送失敗
                sendSuccess = false;

                lastException = smtpEx;
                // 處理SMTP特定錯誤
                HandleSmtpException(smtpEx);
            }
            catch (WebException webEx)
            {
                // 發送失敗
                sendSuccess = false;

                lastException = webEx;
                // 網路連線問題
                HandleNetworkException(webEx);
            }
            catch (Exception ex)
            {
                // 發送失敗
                sendSuccess = false;

                lastException = ex;
                // 其他錯誤
                _logger.LogError(ex, "[func::SendMail126] An unexpected error occurred while sending email.");
            }
            finally
            {
                sw.Stop();
                string status = sendSuccess ? "SUCCESS" : "FAILED";
                string errorInfo = lastException != null ? $"[ERROR: {lastException.GetType().Name} {lastException.Message}]" : "";

                string loggerLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [126 MAIL] [TO:{MailMessageMain.ToMail}] [{status}] [Elapsed:{sw.ElapsedMilliseconds}ms] {errorInfo}";

                Console.WriteLine(loggerLine);
                _logger.LogInformation(loggerLine);

                if (lastException != null)
                {
                    _logger.LogError(lastException, "[Exception] Email sending failure details");
                }

                MailMessage126.Dispose();
            }

            return sendSuccess;
        }
        private void SendMail163()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //163=================================================================================== 
            MailMessage163 = new MailMessage();
            MailMessage163.From = new System.Net.Mail.MailAddress(SenderEmailAccount.FromMailAddress);
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
             

            MailMessage163.Sender = new MailAddress(SenderEmailAccount.FromMailAddress);
            try
            {
                SmtpClient client = new SmtpClient();
                client.Host = SenderEmailAccount.SenderServerHost; //"smtp.163.com";//使用163的SMTP伺服器傳送郵件
                client.UseDefaultCredentials = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                string senderUserName = SenderEmailAccount.SenderUserName;   // _SendMailInfo.SenderUserName.Split('@')[0];

                if (SenderEmailAccount.EnablePasswordAuthentication)
                {
                    //local machine pc Network Credential 網絡憑證
                    System.Net.NetworkCredential nc = new System.Net.NetworkCredential(senderUserName, SenderEmailAccount.SenderUserPassword);
                    client.Credentials = nc.GetCredential(SenderEmailAccount.SenderServerHost, SenderEmailAccount.SenderServerHostPort, "NTLM");
                }
                else
                {
                    client.Credentials = new System.Net.NetworkCredential(senderUserName, SenderEmailAccount.SenderUserPassword);
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

        private void SendMailQQ()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //QQ 
            MailMessageQQ.From = new System.Net.Mail.MailAddress(SenderEmailAccount.FromMailAddress);
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
            MailMessageQQ.Sender = new MailAddress(SenderEmailAccount.FromMailAddress);
            try
            {
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = SenderEmailAccount.SenderServerHost;
                client.UseDefaultCredentials = true;
                client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential(SenderEmailAccount.SenderUserName, SenderEmailAccount.SenderUserPassword); // userName = abc pass = ***   abc@mail.com

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

        private bool SendGmail()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            bool sendSuccess = false;
            Exception lastException = null;

            //GMAIL
            MailMessageGmail = new MailMessage();
            //default
            MailMessageGmail.To.Add(MailMessageMain.ToMail);
#if DEBUG
            MailMessageGmail.Bcc.Add("caihaili82@gmail.com");
#endif
            MailMessageGmail.From = new MailAddress(SenderEmailAccount.FromMailAddress);
            MailMessageGmail.Subject = MailMessageMain.Subject;
            MailMessageGmail.Body = MailMessageMain.EmailBody;
            MailMessageGmail.IsBodyHtml = true;
            MailMessageGmail.SubjectEncoding = System.Text.Encoding.UTF8;
            MailMessageGmail.BodyEncoding = System.Text.Encoding.UTF8;
            MailMessageGmail.IsBodyHtml = true;
            MailMessageGmail.Priority = MailPriority.High;
            MailMessageGmail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
            MailMessageGmail.Sender = new MailAddress(SenderEmailAccount.FromMailAddress);

            try
            {
                if (MailMessageGmail != null)
                {
                    SmtpClient smtpClient = new SmtpClient();
                    smtpClient.Host = SenderEmailAccount.SenderServerHost;
                    smtpClient.Port = SenderEmailAccount.SenderServerHostPort;

                    smtpClient.UseDefaultCredentials = true; //如果使用預設憑證，則為 true；否則為 false

                    smtpClient.EnableSsl = SenderEmailAccount.EnableSSL;
                    SecureString SecureStringOfSenderUserPassword = SecureStringConverter(SenderEmailAccount.SenderUserPassword);
                    if (SenderEmailAccount.EnablePasswordAuthentication)
                    {
                        System.Net.NetworkCredential nc = new System.Net.NetworkCredential(SenderEmailAccount.SenderUserName, SenderEmailAccount.SenderUserPassword);
                        //NTLM: Secure Password Authentication in Microsoft Outlook Express
                        smtpClient.Credentials = nc.GetCredential(smtpClient.Host, smtpClient.Port, "NTLM");
                    }
                    else
                    {
                        smtpClient.Credentials = new System.Net.NetworkCredential(SenderEmailAccount.SenderUserName, SecureStringOfSenderUserPassword, SenderEmailAccount.SenderServerHost);
                    }
                    smtpClient.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                    smtpClient.Send(MailMessageGmail);
                    sendSuccess = true;
                }
            }
            catch (SmtpException smtpEx)
            {
                // 發送失敗
                sendSuccess = false;

                lastException = smtpEx;
                // 处理SMTP特定错误
                HandleSmtpException(smtpEx);
            }
            catch (WebException webEx)
            {
                // 發送失敗
                sendSuccess = false;

                lastException = webEx;
                // 网络连接问题
                HandleNetworkException(webEx);
            }
            catch (Exception ex)
            {
                // 發送失敗
                sendSuccess = false;

                string logerr = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [GMAIL] [TO :{1} [EXCEPTION]:{2}]", DateTime.Now, MailMessageMain.ToMail, ex.Message);
                _logger.LogError(ex, "[func::SendGMail] An unexpected error occurred while sending email.");
            }
            finally
            {
                sw.Stop();
                string status = sendSuccess ? "SUCCESS" : "FAILED";
                string errorInfo = lastException != null ? $"[ERROR: {lastException.GetType().Name}]" : "";

                string loggerLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [126 MAIL] [TO:{MailMessageMain.ToMail}] [{status}] [Elapsed:{sw.ElapsedMilliseconds}ms] {errorInfo}";

                Console.WriteLine(loggerLine);
                _logger.LogInformation(loggerLine);

                if (lastException != null)
                {
                    _logger.LogError(lastException, "邮件发送失败详情");
                }

                MailMessage126.Dispose();
            }
            return sendSuccess;
        }

        private bool SendMail()
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            bool sendSuccess = false;
            Exception lastException = null;

            try
            {
                // 初始化邮件消息
                MailMessage126.From = new MailAddress(SenderEmailAccount.FromMailAddress);
                MailMessage126.To.Add(MailMessageMain.ToMail);
#if DEBUG
                MailMessage126.Bcc.Add("caihaili82@gmail.com");
#endif
                MailMessage126.Subject = MailMessageMain.Subject;
                MailMessage126.Body = MailMessageMain.EmailBody;
                MailMessage126.IsBodyHtml = true;
                MailMessage126.SubjectEncoding = Encoding.UTF8;
                MailMessage126.BodyEncoding = Encoding.UTF8;
                MailMessage126.Priority = MailPriority.High;
                MailMessage126.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
                MailMessage126.Sender = new MailAddress(SenderEmailAccount.FromMailAddress);

                using SmtpClient client = new SmtpClient();
                // 设置SMTP服务器信息
                client.Host = SenderEmailAccount.SenderServerHost;
                client.Port = SenderEmailAccount.SenderServerHostPort; // 显式指定端口
                client.EnableSsl = SenderEmailAccount.EnableSSL;
                client.Timeout = 10000; // 设置10秒超时
                client.UseDefaultCredentials = false;

                // 身份验证
                client.Credentials = new NetworkCredential(
                    SenderEmailAccount.SenderUserName,
                    SenderEmailAccount.SenderUserPassword
                );

                // 发送邮件
                client.Send(MailMessage126);
                sendSuccess = true;
            }
            catch (SmtpException smtpEx)
            {
                lastException = smtpEx;
                // 处理SMTP特定错误
                HandleSmtpException(smtpEx);
            }
            catch (WebException webEx)
            {
                lastException = webEx;
                // 网络连接问题
                HandleNetworkException(webEx);
            }
            catch (Exception ex)
            {
                lastException = ex;
                // 其他错误
                _logger.LogError(ex, "[func::SendMail] An unexpected error occurred while sending email.");
            }
            finally
            {
                sw.Stop();
                string status = sendSuccess ? "SUCCESS" : "FAILED";
                string errorInfo = lastException != null ? $"[ERROR: {lastException.GetType().Name}]" : "";

                string loggerLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [126 MAIL] [TO:{MailMessageMain.ToMail}] [{status}] [Elapsed:{sw.ElapsedMilliseconds}ms] {errorInfo}";

                Console.WriteLine(loggerLine);
                _logger.LogInformation(loggerLine);

                if (lastException != null)
                { 
                    _logger.LogError(lastException, "[func::SendMail126] Email sending failure details.");
                }

                MailMessage126.Dispose();
            } 
            return sendSuccess;
        }

        /// <summary>
        /// 異步發送郵件 Sending emails asynchronously
        /// </summary>
        /// <returns></returns>
        public async Task<bool> SendMailsync(string toMail, string subject, string emailBody)
        {
            MailMessageMain = new MailMessageMain { ToMail = toMail, Subject = subject, EmailBody = emailBody };

            bool sendSuccess = false;

            string senderOfCompany = SenderEmailAccount.SenderOfCompany.ToLower();

            await Task.Run(() =>
            {
                //Console.Write($"[SendMail] [SenderOfCompany: {senderOfCompany}]\n\n[ToMail: {toMail}]\n \n[Subject: {subject}]\n\n [EmailBody: {emailBody}]");
                Console.Write($"[SendMail] [SenderOfCompany: {senderOfCompany}]\n\n[ToMail: {toMail}]\n \n[Subject: {subject}]");
            });

            try
            {
                switch (senderOfCompany)
                {
                    case "126.com":
                        sendSuccess = SendMail126();
                        break;
                    case "163.com":
                        SendMail163();
                        break;
                    case "qq.com":
                        SendMailQQ();
                        break;
                    case "gmail.com":
                        sendSuccess = SendGmail();
                        break;
                    default:
                        sendSuccess = SendMail();
                        break;
                }
                return sendSuccess;
            }
            catch
            {
                return sendSuccess;
            }
        }

        // 處理SMTP特定錯誤
        private void HandleSmtpException(SmtpException ex)
        {
            switch (ex.StatusCode)
            {
                case SmtpStatusCode.GeneralFailure:
                    _logger.LogError(ex, $"[SmtpException] SMTP server connection failed, please check the server address and port {ex.StatusCode} {ex.Message}");
                    break;
                case SmtpStatusCode.ClientNotPermitted:
                case SmtpStatusCode.MustIssueStartTlsFirst:
                    _logger.LogError(ex, "SMTP authentication failed, please check your username and password");
                    break;
                case SmtpStatusCode.MailboxBusy:
                case SmtpStatusCode.InsufficientStorage:
                    _logger.LogWarning(ex, "The SMTP server is temporarily unavailable. Please try again later.");
                    break;
                default:
                    _logger.LogError(ex, $"SMTP ERROR: {ex.StatusCode}");
                    break;
            }
        }

        // 處理網路連線錯誤
        private void HandleNetworkException(WebException ex)
        {
            switch (ex.Status)
            {
                case WebExceptionStatus.ConnectFailure:
                    _logger.LogError(ex, "无法连接到SMTP服务器，请检查网络连接");
                    break;
                case WebExceptionStatus.Timeout:
                    _logger.LogError(ex, "连接SMTP服务器超时，请检查服务器状态或增加超时时间");
                    break;
                case WebExceptionStatus.NameResolutionFailure:
                    _logger.LogError(ex, "无法解析SMTP服务器地址，请检查主机名是否正确");
                    break;
                default:
                    _logger.LogError(ex, $"网络错误: {ex.Status}");
                    break;
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
     
}

