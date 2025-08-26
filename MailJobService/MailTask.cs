using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;   
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection; 
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Text.Encodings.Web;
using System.Diagnostics;
using System.Text;

namespace MailJobService
{
    public class EmailTask
    {
        public static async Task RunAsync(string[] argsSendToMailArray,string subject, MailTemplateEnum mailTemplateEnum, string LanguageCode,string callBackUrlEncode,string attachPath)
        {
            // 创建ServiceCollection
            var services = new ServiceCollection();
            ConfigureServices(services);
             
            // 创建ServiceProvider
            var serviceProvider = services.BuildServiceProvider();
            var logger = serviceProvider.GetRequiredService<ILogger<Program>>();

            try
            {
                logger.LogInformation("[App is start running ]"); 
                //serviceProvider.GetService<App>().Run(args);
                await serviceProvider.GetService<App>().Run(argsSendToMailArray, subject,mailTemplateEnum, LanguageCode, callBackUrlEncode, attachPath); 
            }
            catch (Exception ex)
            {
                logger.LogInformation(ex, "[APP RUNNING EXCEPTION ERROR OCCURED ]");
            } 
        }
        public static IConfigurationRoot ReadFromAppSettings()
        {
            return new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", false)
                .AddEnvironmentVariables()
                .Build();
        }
         
        // This method gets called by the runtime. Use this method to add services to the container.
        public static void ConfigureServices(IServiceCollection services)
        {
            // 创建 appsettings
            var Configuration = ReadFromAppSettings(); 
             
            services.AddOptions();

            services.Configure<SendMailInfo>(Configuration.GetSection("SendMailInfo"));
            services.Configure<GmailInfo>(Configuration.GetSection("GmailInfo"));

            services.AddScoped<SendMailInfo>();
            services.AddScoped<GmailInfo>();
            services.AddTransient<App>();


            // 配置日志
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.AddDebug();
            });
        }
    }
    /// <summary>
    /// 【独立传入SendMailInfo值】或者【版本】使用独立配置文件 EmailServiceConfig.json 也可以选择对象服务，从Appsetting。json获取。
    /// 经测试，OK。 GMAIL需要开启【阻止某些不安全设备或应用登录google账号。】
    /// </summary>
    public class EmailApp
    { 
        public EmailApp(SendMailInfo sendMailInfo, GmailInfo gmailInfo)
        {  
            SendMailInfo = sendMailInfo;
            GmailInfo = gmailInfo;
        }
        public SendMailInfo SendMailInfo { get; set; }
        public GmailInfo GmailInfo { get; set; }

        public async Task Run(string[] args, string subject, string mailTemplateEnum, string languageCode, string callBackUrlEncode, string attachPath)
        {
            #region BEGIN
            //============================================================================================================================== 
            //Stopwatch sw = new Stopwatch();
            //sw.Start();

            foreach (var item in args)
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
                    string bodyInfo = GetMailTemplate(mailTemplateEnum, languageCode);
                    if (!string.IsNullOrEmpty(bodyInfo))
                    {
                        bodyInfo = bodyInfo.Replace("{callbackurl}", callBackUrlEncode);
                    }
                    string subjectInfo = string.Empty;
                    
                    if (!string.IsNullOrEmpty(subject))
                    {
                        subjectInfo = subject;
                    }
                    if (subjectInfo.Length<=1)
                    {
                        string htmlContent = GetHtmlText(bodyInfo);
                        int cutLenght = htmlContent.Length > 20 ? 20 : htmlContent.Length;
                        subjectInfo = htmlContent.Substring(0, cutLenght);
                    }

                    if (languageCode.ToLower() == "zh-cn")
                    {
                        EmailHelper email = new EmailHelper(SendMailInfo, toMailAddress, subjectInfo, bodyInfo);

                        email.AddAttachments(attachPath);
                        switch (SendMailInfo.SenderOfCompany)
                        {
                            case "163":
                                email.SendMail163();
                                //CommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [163 MAIL] [TO :{1} Form:{2} Subject:{3}]",DateTime.Now, SendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                            case "126":
                                email.SendMail126();
                                //CommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [126 MAIL] [TO :{1} Form:{2} Subject:{3}]", DateTime.Now, SendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                            case "QQ":
                                email.SendMailQQ();
                                //CommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [QQ MAIL] [TO :{1} Form:{2} Subject:{3}]", DateTime.Now, SendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                            case "GOOGLE":
                                email.SendMailGoogle();
                                //CommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [GOOGLE GMAIL] [TO :{1} Form:{2} Subject:{3}]", DateTime.Now, SendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                            default:
                                email.SendMail();
                                //CommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}] [DETAULT MAIL] [TO :{1} Form:{2} Subject:{3}]", DateTime.Now, SendMailInfo.FromMailAddress, toMailAddress, subjectInfo));
                                break;
                        }
                        Console.WriteLine("[{0:yyyy-MM-dd HH:mm:ss}] Sending Email （{1}）......", DateTime.Now, toMailAddress);
                    }
                    else
                    {
                        GmailHelper gmail = new GmailHelper(GmailInfo, toMailAddress, subjectInfo, bodyInfo);

                        gmail.AddAttachments(attachPath);

                        gmail.SendMailGoogle();

                        Console.WriteLine("[{0:yyyy-MM-dd HH:mm:ss}][EmailApp Run] Sending GEmail （{1}）......", DateTime.Now, toMailAddress);

                        MailCommonBase.OperateDateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][EmailApp Run] TO :{1} Form:{2} Subject:{3}", DateTime.Now, GmailInfo.senderMailAddress, toMailAddress, subjectInfo));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                #endregion 
            }
             
            //await Task.CompletedTask; 
            
            //end
            //string time = sw.Elapsed + "-(" + sw.Elapsed.Seconds + " seconds)";
            //sw.Stop();
            //Console.WriteLine(">>> Elapsed Time = {0}", sw.Elapsed);

            #endregion END 
        }
        /// <summary>
        /// MailTemplateEnum 常量类型 改用字符串类型 例如：模板 ForgetPassword_zh-HK.html  常量定义是：ForgetPassword
        /// </summary>
        /// <param name="mailTemplateEnum"></param>
        /// <param name="LanguageCode"></param>
        /// <returns></returns>
        public string GetMailTemplate(string  mailTemplateEnum, string LanguageCode)
        { 
            string fileName = string.Format("{0}_{1}.html", mailTemplateEnum, LanguageCode);
            string content = fileName;
            string pathFileTemplate = Path.Combine(Directory.GetCurrentDirectory(), "MailTemplate", fileName);

            if (!File.Exists(pathFileTemplate))
            {
                Console.WriteLine("[GetMailTemplate] FILE NOT EXIST [{0}] [CHECK FOLDER (MailTemplate)]", fileName);
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
    }
}

