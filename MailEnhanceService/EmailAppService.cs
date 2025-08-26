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
    /// 【獨立傳入 Sender Of Email Account List 】或【版本】使用獨立設定檔 EmailServiceConfig.json 也可以選擇物件服務，從Appsetting。 json獲取。 
    /// 經測試，OK。 GMAIL需要開啟【阻止某些不安全設備或應用程式登入google帳號。 】
    /// </summary>
    public class EmailAppService
    { 
        public readonly ILogger<EmailAppService> _logger;

        private EmailEnhanceHelper emailEnhanceHelper { get; set; }

        private IList<SenderEmailAccount> SenderEmailAccountList { get; set; }

        /// <summary>
        /// EmailAppService 构造函数
        /// </summary>
        /// <param name="mailComId">如果存在MAinComId 則從本地調取發郵件賬號列表</param>
        /// <param name="emailEnhance"></param>
        /// <param name="senderEmailAccountList"></param>
        /// <param name="logger"></param>
        public EmailAppService(string mainComId, EmailEnhanceHelper emailEnhance, IList<SenderEmailAccount> senderEmailAccountList, ILogger<EmailAppService> logger)
        {
            _logger = logger; 
            emailEnhanceHelper = emailEnhance;

            //沒有指明 MainComId,則使用系統Aspsettings.json 中的配置
            if (string.IsNullOrEmpty(mainComId))
            {
                _logger.LogInformation($"Because mainComId is an empty string, the email account list is retrieved from the local configuration. {nameof(mainComId)}");
                SenderEmailAccountList = senderEmailAccountList;
            }
            else
            {
                bool isFromMainComId = ToInstanceFromLocalConfigJson(mainComId,out IList<SenderEmailAccount> SenderList); 
                if (isFromMainComId)
                {
                    SenderEmailAccountList = SenderList;
                    _logger.LogInformation($"The Sender Email Account List Is From MainCom(Config Folder) {mainComId}");
                }
                else
                {
                    //獲取不到配置文件，則使用傳入的 senderEmailAccountList
                    SenderEmailAccountList = senderEmailAccountList;
                }
            }
               
            _logger.LogInformation("Starting EmailAppService...");
        }

        /// <summary>
        /// 啟動 EmailAppService 服務
        /// 首先啟動服務後再調用發郵件函數
        /// </summary>
        /// <returns></returns> 
        public static EmailAppService StartUpEmailAppService(string? mainComId)
        {
            if (string.IsNullOrEmpty(mainComId))
            {
                mainComId = string.Empty;
            }

            // 建立服務集合並配置服務
            var services = new ServiceCollection();
            ServiceConfigurator.ConfigureServices(services, mainComId);

            // 建構服務提供者
            var serviceProvider = services.BuildServiceProvider();
             
            // 获取邮件服务
            using var scope = serviceProvider.CreateScope();
            var emailAppService = scope.ServiceProvider.GetRequiredService<EmailAppService>();
            return emailAppService;
        }

        /// <summary>
        /// 傳入郵件列表，逐一發送郵件
        /// 具體用法參考 Program.cs 或 MailJobService/Program.cs.md
        /// </summary>
        /// <param name="mailToList">email 字符串数组</param>
        /// <param name="subjectIfHave">如果沒有主題，則會使用 內容的純文本前20字作為主題。</param>
        /// <param name="mailTemplateEnum">不使用模板，必須 MailTemplateEnum.NO_TEMPLATE</param>
        /// <param name="bodyRawContentIfHave">如果指定了郵件模版，則會使用指定的模板内容，而不會例會傳入參數 bodyRawContent 的內容</param>
        /// <param name="languageCode">使用什麼語言版本的模板 en-US;zh-HK;zh-CN</param>
        /// <param name="callBackUrlEncode">回調的URL</param>
        /// <param name="attachedFilesIfHave">文件路徑的字符串數組</param>
        /// <returns></returns>
        public async Task<bool> RunAsync(string[] mailToList, string subjectIfHave, MailTemplateEnum mailTemplateEnum, string bodyRawContentIfHave, string languageCode, string callBackUrlEncode, string[]? attachedFilesIfHave=null)
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
                        bodyRawContentIfHave = GetMailTemplate(mailTemplateEnum, languageCode);
                    }

                    if (!string.IsNullOrEmpty(bodyRawContentIfHave))
                    {
                        bodyRawContentIfHave = bodyRawContentIfHave.Replace("{callbackurl}", callBackUrlEncode);
                    }
                    string subjectInfo = string.Empty;

                    if (!string.IsNullOrEmpty(subjectIfHave))
                    {
                        subjectInfo = subjectIfHave;
                    }
                    else
                    {
                        //如果主題信息為空，則從正文內容中提取前50個字符作為主題
                        string htmlContent = GetHtmlText(bodyRawContentIfHave);
                        int cutLenght = htmlContent.Length > 50 ? 50 : htmlContent.Length;
                        subjectInfo = htmlContent.Substring(0, cutLenght);
                    }

                    //發送郵件， 後續 可在這裡配合設置多個郵件發送賬號，啟動發送失敗輪詢發送郵件，
                    foreach (var item1OfSenderAccount in SenderEmailAccountList)
                    {
                        bool succ = await emailEnhanceHelper.SendMailsync(item1OfSenderAccount,toMailAddress, subjectInfo, bodyRawContentIfHave, attachedFilesIfHave);
                         
                        if (succ)
                        {
                            _logger.LogInformation($"Sending Email the first time（{toMailAddress}） SUCC = {succ}......");
                            break; // 如果發送成功，跳出循環
                        }
                        else
                        {
                            _logger.LogError($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [sender:{item1OfSenderAccount.SenderUserName}] Failed to send email to {toMailAddress}. Trying next sender account...");
                        }
                    } 
                }
                catch (Exception ex)
                {
                    const string logTemplate = "[Exeception] [{Message}]";
                    _logger.LogError(logTemplate, ex.Message);
                }
                finally
                {
                    _logger.LogInformation($">>>Send Mail Elapsed Time = {sw.Elapsed.Seconds} Seconds");
                    sw.Stop();
                }
            }

            return true;

            #endregion END TASK RUN
        }

        /// <summary>
        /// MailTemplateEnum 常量类型 改用字符串类型 例如：模板 ForgetPassword_zh-HK.html  常量定义是：ForgetPassword
        /// 可提供外部調用，根據模板枚舉值和語言代碼獲取對應的郵件模板內容
        /// </summary>
        /// <param name="mailTemplateEnum"></param>
        /// <param name="LanguageCode"></param>
        /// <returns></returns>
        public string GetMailTemplate(MailTemplateEnum mailTemplateEnum, string LanguageCode)
        {
            string basePath = AppContext.BaseDirectory;
             
            string templateFileName = FormatTemplateName(mailTemplateEnum);  
            string fileName = string.Format("{0}_{1}.html", templateFileName, LanguageCode);

            string content = fileName;

            string pathFileTemplate = Path.Combine(basePath, "MailTemplate", fileName);  //Directory.GetCurrentDirectory() 、MailCommonBase.BasePath

            if (!File.Exists(pathFileTemplate))
            {
                _logger.LogInformation("[GetMailTemplate] FILE NOT EXIST [{0}] [CHECK FOLDER (MailTemplate)]", fileName);
                _logger.LogInformation($"[MAilTask.GetMailTemplate][] FILE NOT EXIST [{pathFileTemplate}]");

                // 試試透過HostingEnvironment.ContentRootPath 取得路徑
                var host = Host.CreateDefaultBuilder().Build();
                var hostEnv = host.Services.GetRequiredService<IHostEnvironment>();
                basePath = hostEnv.ContentRootPath;
                pathFileTemplate = Path.Combine(basePath, "MailTemplate", fileName);
                // 再次驗證或不存在則傳回內容為檔案名稱 以提醒後續跟進檔案路徑是否有同步到目前目錄下。
                if (!File.Exists(pathFileTemplate))
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
        /// Get plain text in HTML ref  https://blog.csdn.net/fuzhixin0/article/details/52129253
        /// </summary>
        /// <param name="html">html</param>
        /// <returns>plain text</returns>
        public string GetHtmlText(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }
            html = System.Text.RegularExpressions.Regex.Replace(html, @"<\/*[^<>]*>", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            html = html.Replace("\r\n", "").Replace("\r", "").Replace("&nbsp;", "").Replace(" ", "");
            return html;
        }

        /// <summary>
        /// 將枚舉值轉換為郵件模板文件名稱
        /// </summary>
        /// <param name="template"></param>
        /// <returns></returns>
        public string FormatTemplateName(MailTemplateEnum template)
        {
            // 將枚舉值轉換為字串
            string templateName = template.ToString();
            // 按下劃線分割為單字數組
            string[] words = templateName.Split('_');
            // 建構結果字串
            StringBuilder result = new StringBuilder();
            foreach (string word in words)
            {
                if (!string.IsNullOrEmpty(word))
                {
                    // 首字母大楷，其餘细楷
                    result.Append(char.ToUpper(word[0]) + word.Substring(1).ToLower());
                }
            }
            return result.ToString();
        }

        /// <summary>
        /// 如果有配置自定義的 郵箱發送賬號列表文件（./Config/SenderEmailAccountList.json）
        /// </summary>
        /// <param name="mainComId">一個定義的前綴，預設架構是 多個公司同時使用此軟件的情況。如果單一公司則傳入 string.empty</param>
        /// <param name="senderEmailAccountList">如果引用的 senderEmailAccountList 發送賬號列表為空，則嘗試本地配置文件夾 ./config</param>
        /// <returns>是否引用本地config文件夾的配置</returns>
        public static bool ToInstanceFromLocalConfigJson(string mainComId, out IList<SenderEmailAccount> senderListOfMainCom)
        {
            senderListOfMainCom = new List<SenderEmailAccount>();

            string configPath = Path.GetFullPath("./Config");
            string configFileName = $"SenderEmailAccountList.Json"; 
            if(!string.IsNullOrEmpty(mainComId))
            { 
                configFileName = $"SenderEmailAccountList_{mainComId}.json";
            } 
            string pathfileName = Path.Combine(configPath, configFileName);

            if(!File.Exists(pathfileName))
            {
                // If the file does not exist, returns an empty list. 
                return false;
            }

            string jsonOfMailSenderList = MailCommonBase.ReadConfigJson(pathfileName);
            if (!string.IsNullOrEmpty(jsonOfMailSenderList))
            {
                senderListOfMainCom = JsonConvert.DeserializeObject<IList<SenderEmailAccount>>(jsonOfMailSenderList);

                if (senderListOfMainCom!=null && senderListOfMainCom.Count > 0)
                { 
                    return true; // If the list is empty after deserialization, return false
                } 
            }
            return false;
        }
    }
}

