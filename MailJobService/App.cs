using System;
using System.Collections.Generic;
using System.Text;
using System.Linq; 

using System.IO;
using System.Net.Mail;
using System.Net;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options; 
using System.Threading.Tasks; 
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;  
using System.Threading;  
using System.Diagnostics;
using System.Globalization;  

namespace MailJobService
{
    /// <summary>
    /// FOLDER In MailTemplate
    /// </summary>
    public enum MailTemplateEnum
    {
        Register = 0,
        ForgetPassword = 1,
        EMAIL_MARKETING = 2,
        PrivacyContent =3
    }
    public class App
    {
        private ILogger<App> _logger;
        private SendMailInfo _sendMailInfo;
        private GmailInfo _gmailInfo;
         
        public App(IOptions<SendMailInfo> sendMailInfo, IOptions<GmailInfo> gmailInfo, ILogger<App> logger)
        { 
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));

            _sendMailInfo = sendMailInfo.Value;

            _gmailInfo = gmailInfo.Value; 
        }
        
        public async Task Run(string[] args,string subject, MailTemplateEnum mailTemplateEnum, string LanguageCode,string callBackUrlEncode,string attachPath)
        { 
            //Console.Clear();
            Console.OutputEncoding = Encoding.UTF8; 
             
            #region BEGIN
            //==============================================================================================================================
               
            Stopwatch sw = new Stopwatch();
            sw.Start();

            foreach(var item in args)
            {
                string waitToSend = item;
                 
                if (!MailCommonBase.IsValidEmail(waitToSend))
                {
                    continue;
                }

                #region EmailHelper
                try
                {
                    string toMailAddress = waitToSend; 
                    await Task.Delay(1);
                    string bodyInfo = GetMailTemplate(mailTemplateEnum, LanguageCode);
                    if(!string.IsNullOrEmpty(bodyInfo))
                    {
                        bodyInfo = bodyInfo.Replace("{callbackurl}", callBackUrlEncode);
                    }
                    string subjectInfo = string.Empty;

                    if (!string.IsNullOrEmpty(subject))
                    {
                        subjectInfo = subject;
                    }

                    if (subjectInfo.Length <= 1)
                    {
                        string htmlContent = GetHtmlText(bodyInfo);
                        int cutLenght = htmlContent.Length > 20 ? 20 : htmlContent.Length;
                        subjectInfo = htmlContent.Substring(0, cutLenght);
                    }

                    if (LanguageCode.ToLower() == "zh-cn")
                                                                                                                                                                         {
                        EmailHelper email = new EmailHelper(_sendMailInfo, toMailAddress, subjectInfo, bodyInfo);
                        email.AddAttachments(attachPath);

                        switch (_sendMailInfo.SenderOfCompany)
                        {
                            case "163":
                                email.SendMail163();
                                MailCommonBase.OperateDateLoger(string.Format("[163 MAIL] [TO :{0} Form:{1} Subject:{2}]", _sendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                            case "126":
                                email.SendMail126();
                                MailCommonBase.OperateDateLoger(string.Format("[126 MAIL] [TO :{0} Form:{1} Subject:{2}]", _sendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                            case "QQ":
                                email.SendMailQQ();
                                MailCommonBase.OperateDateLoger(string.Format("[QQ MAIL] [TO :{0} Form:{1} Subject:{2}]", _sendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                            default:
                                email.SendMail();
                                MailCommonBase.OperateDateLoger(string.Format("[DETAULT MAIL] [TO :{0} Form:{1} Subject:{2}]", _sendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                        }
                        _logger.LogInformation("Sending Email......");
                    }else
                    {

                        GmailHelper gmail = new GmailHelper(_gmailInfo, toMailAddress, subjectInfo, bodyInfo);

                        gmail.AddAttachments(attachPath);

                        gmail.SendMailGoogle();

                        _logger.LogInformation("Sending Gmail......");

                        MailCommonBase.OperateDateLoger(string.Format("TO :{0} Form:{1} Subject:{2}", _gmailInfo.senderMailAddress, toMailAddress, subjectInfo));
                    } 
                }
                catch (Exception ex)
                {
                    _logger.LogInformation("[EXCEPTION] {0}", ex.ToString());

                    MailCommonBase.OperateDateLoger(string.Format("TO :{0} Form:{1} [EXCEPTION] :{2}", waitToSend, _gmailInfo.senderMailAddress, ex.Message));
                }
                #endregion 
            }
            //end
            string time = sw.Elapsed + "-(" + sw.Elapsed.Seconds + " seconds)";
            sw.Stop();
             
            _logger.LogInformation(">>> Elapsed Time = {0}", sw.Elapsed);
            _logger.LogInformation("[SendMailInfo] App.Run() Finished!");
            
            //await Task.CompletedTask; 
            #endregion END 
        }

        public string GetMailTemplate(MailTemplateEnum mailTemplateEnum,string LanguageCode)
        {
            string content = "  ";
            string fileName = string.Format("{0}_{1}.html", mailTemplateEnum.ToString(), LanguageCode);
            string pathFileTemplate = Path.Combine(Directory.GetCurrentDirectory(), "MailTemplate", fileName);

            if(!File.Exists(pathFileTemplate))
            {
                _logger.LogInformation("MAIL TEMPLATE PATH IS NOT EXIST [{0}]", pathFileTemplate); 
                return content;
            }
            try
            {
                content = File.ReadAllText(pathFileTemplate, Encoding.UTF8);

                return content;
            }
            catch
            { 
                return content;
            }
        }

        /// <summary>
        /// 获取html中纯文本 ref  https://blog.csdn.net/fuzhixin0/article/details/52129253
        /// </summary>
        /// <param name="html">html</param>
        /// <returns>纯文本</returns>
        public string GetHtmlText(string html)
        {
            html = System.Text.RegularExpressions.Regex.Replace(html, @"<\/*[^<>]*>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            html = html.Replace("\r\n", "").Replace("\r", "").Replace("&nbsp;", "").Replace(" ", "");
            return html;
        } 
    }
}
