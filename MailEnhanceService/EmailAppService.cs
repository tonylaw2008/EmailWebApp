using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using Microsoft.Extensions.Configuration; 
using Microsoft.Extensions.DependencyInjection; 
using Microsoft.Extensions.Logging;
using Microsoft.Extensions;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Text.Encodings.Web;
using System.Diagnostics;
using System.Text;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Builder;
using System.Reflection;
using Microsoft.Extensions.Hosting;

namespace MailEnhanceService
{

    /// <summary>
    /// 【独立传入SendMailInfo值】或者【版本】使用独立配置文件 EmailServiceConfig.json 也可以选择对象服务，从Appsetting。json获取。
    /// 经测试，OK。 GMAIL需要开启【阻止某些不安全设备或应用登录google账号。】
    /// </summary>
    public class EmailAppService
    {
        private readonly ILogger<EmailAppService> _logger;
        private IList<SenderEmailAccount> SenderEmailAccountList { get; set; }
        //默認的發件箱賬戶信息為列表第一個賬戶信息
        private SenderEmailAccount SenderEmailAccount { get; set; }

        private EmailEnhanceHelper EmailEnhanceHelper { get; set; }

        public EmailAppService(EmailEnhanceHelper emailEnhance, IList<SenderEmailAccount> senderEmailAccountList, ILogger<EmailAppService> logger)
        {
            EmailEnhanceHelper = emailEnhance;
            SenderEmailAccountList = senderEmailAccountList;
            SenderEmailAccount = SenderEmailAccountList[0];   //保證絕對的aspsetting.json配置,則肯定不會為null
            _logger = logger;
        }

        /// <summary>
        /// 启动 EmailAppService 服务
        /// 首先啟動服務後再調用發郵件函數
        /// </summary>
        /// <returns></returns>
        public static EmailAppService StartUpEmailAppService()
        {
            // 创建服务集合
            var services = new ServiceCollection();
            ServiceConfigurator.ConfigureServices(services);

            // 构建服务提供者
            var serviceProvider = services.BuildServiceProvider();

            // 获取邮件服务
            using var scope = serviceProvider.CreateScope();
            var emailAppService = scope.ServiceProvider.GetRequiredService<EmailAppService>();
            return emailAppService;
        }
        public async Task<bool> RunAsync(string[] mailToList, string subject, MailTemplateEnum mailTemplateEnum, string bodyRawContent, string languageCode, string callBackUrlEncode, string attachPath)
        {
            #region BEGIN TASK RUN
            //============================================================================================================================== 

            foreach (var item in mailToList)
            {
                Stopwatch sw = new Stopwatch();
                sw.Start();
                string waitToSend = item;

                if (!MailCommonBase.IsValidEmail(waitToSend))
                {
                    continue;
                }

                try
                {
                    string toMailAddress = waitToSend;

                    if (mailTemplateEnum != MailTemplateEnum.NO_TEMPLATE)
                    {
                        bodyRawContent = GetMailTemplate(mailTemplateEnum.ToString(), languageCode);
                    }

                    if (!string.IsNullOrEmpty(bodyRawContent))
                    {
                        bodyRawContent = bodyRawContent.Replace("{callbackurl}", callBackUrlEncode);
                    }
                    string subjectInfo = string.Empty;

                    if (!string.IsNullOrEmpty(subject))
                    {
                        subjectInfo = subject;
                    }

                    //如果主題信息為空，則從正文內容中提取前20個字符作為主題
                    if (subjectInfo.Length <= 1)
                    {
                        string htmlContent = GetHtmlText(bodyRawContent);
                        int cutLenght = htmlContent.Length > 20 ? 20 : htmlContent.Length;
                        subjectInfo = htmlContent.Substring(0, cutLenght);
                    }
                    //發送郵件， 後續 可在這裡配合設置多個郵件發送賬號，啟動發送失敗輪詢發送郵件，
                    bool succ = await EmailEnhanceHelper.SendMailsync(toMailAddress, subjectInfo, bodyRawContent);
                    Console.WriteLine("[{0:yyyy-MM-dd HH:mm:ss}] Sending Email （{1}） SUCC = {2}......", DateTime.Now, toMailAddress, succ);

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    Console.WriteLine(">>>Send Mail Elapsed Time = {0}", sw.Elapsed.Seconds);
                    sw.Stop();
                }
            }

            return true;

            #endregion END TASK RUN
        }

        /// <summary>
        /// MailTemplateEnum 常量类型 改用字符串类型 例如：模板 ForgetPassword_zh-HK.html  常量定义是：ForgetPassword
        /// </summary>
        /// <param name="mailTemplateEnum"></param>
        /// <param name="LanguageCode"></param>
        /// <returns></returns>
        public string GetMailTemplate(string mailTemplateEnum, string LanguageCode)
        {
            string fileName = string.Format("{0}_{1}.html", mailTemplateEnum, LanguageCode);
            string content = fileName;

            string pathFileTemplate = Path.Combine(MailCommonBase.BasePath, "MailTemplate", fileName);  //Directory.GetCurrentDirectory()

            if (!File.Exists(pathFileTemplate))
            {
                _logger.LogInformation("[GetMailTemplate] FILE NOT EXIST [{0}] [CHECK FOLDER (MailTemplate)]", fileName);
                MailCommonBase.OperateLoger($"[MAilTask.GetMailTemplate][] FILE NOT EXIST [{pathFileTemplate}]");
                return content;
            }
            Console.WriteLine("[GetMailTemplate] FILE EXIST [{0}]", pathFileTemplate);
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

        //  取得預設的寄件匣帳戶訊息list    
        public IList<SenderEmailAccount> GetSenderEmailAccountList()
        {
            if (SenderEmailAccountList == null || SenderEmailAccountList.Count == 0)
            {
                _logger.LogError("SenderEmailAccountList is empty or null.");
                throw new InvalidOperationException("SenderEmailAccountList is empty or null.");
            }
            return SenderEmailAccountList; // 返回默认的发件箱账户信息
        }
    }
}

