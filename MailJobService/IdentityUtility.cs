using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;  

namespace MailJobService
{
    public partial class IdentityHelper
    {
        public static void SendComfirmRegisterEmail(SendMailInfo sendMailInfo,string LanguageCode, string userName,string toEmail,string contentRootPath, string callbackurl)
        {
            try
            { 
                string senderServerHost = sendMailInfo.SenderServerHost; 
                string toMailAddress = toEmail;
                string fromMailAddress = sendMailInfo.FromMailAddress;
                string subjectInfo = $"Hello,{userName}";
                string templateFile = Path.Combine(contentRootPath, "MailTemplate", $"Register_{LanguageCode}.html");
                string bodyInfo = MailTemplateText(templateFile).Replace("{callbackurl}", callbackurl);
                string mailUsername = sendMailInfo.SenderUserName;
                string mailPassword = sendMailInfo.SenderUserPassword;
                int mailPort = sendMailInfo.SenderServerHostPort;
                 
                string attachPath = ""; 
                EmailHelper email = new EmailHelper(sendMailInfo, toMailAddress, subjectInfo, bodyInfo);
                email.AddAttachments(attachPath);  
                email.SendEmailSync();

            }
            catch (Exception ex)
            { 
                Console.WriteLine(ex.Message);
                MailCommonBase.OperateDateLoger(string.Format("[SendComfirmRegisterEmail] [TO :{0} Form:{1}] [ EXCEPTION:{2}]", toEmail, sendMailInfo.FromMailAddress, ex.Message));
            }
        }
        public static string MailTemplateText(string filepath)
        {
            if (File.Exists(filepath))
            {
                string content = File.ReadAllText(filepath, Encoding.UTF8);
                return content;
            }
            else
            {
                return string.Empty;
            } 
        } 
    }
}
