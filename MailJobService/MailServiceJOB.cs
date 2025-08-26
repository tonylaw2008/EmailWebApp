using System;
using System.Text;
using System.Threading;
using System.Timers; 
using System.Collections.Specialized;
using System.Threading.Tasks;
using Quartz.Impl;
using Quartz;
using Quartz.Logging;
using Quartz.Impl.Matchers;
using static Quartz.MisfireInstruction;
using Microsoft.Extensions.Configuration;
using System.IO;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Net.Http;
using System.Net.Sockets;
using Newtonsoft.Json;

namespace MailJobService
{
    public partial class MailJobServiceQuartz
    { 
        public static async Task RunScheduleMonthlyGlobalProgram(DateTime TaskMonthlyStartDate, int intervalMinutes)
        {
            if (intervalMinutes == 0)
            {
                intervalMinutes = 1;
            }
            //int timesOfTaskRunning = 1; 
            try
            {
                NameValueCollection props = new NameValueCollection
                {
                    { "quartz.serializer.type", "binary" }
                };
                StdSchedulerFactory factory = new StdSchedulerFactory(props);
                IScheduler scheduler = await factory.GetScheduler();

                await scheduler.Start();

                IJobDetail job = JobBuilder.Create<MailJobServiceJOB>()
                    .WithIdentity("MailJobServiceJOB", "GROUP1")
                    .Build();
                ITrigger trigger = TriggerBuilder.Create()
                    .WithIdentity("MailJobServiceTRIGER")
                    .StartAt(TaskMonthlyStartDate)
                    .ForJob("MailJobServiceJOB", "GROUP1")
                    .WithCalendarIntervalSchedule(w => w
                    .WithIntervalInMinutes(intervalMinutes)
                    )  
                    .Build();
                 
                await scheduler.ScheduleJob(job, trigger);
                 
                await Task.Delay(TimeSpan.FromMilliseconds(1000)); 
            }
            catch (SchedulerException se)
            {
                await Console.Error.WriteLineAsync(se.ToString());
            }
        }
        
        public static string CreateDynamicToken(string Scret)
        {
            DateTimeOffset dateTimeOffset = new DateTimeOffset(DateTime.Now);
            long ts1 = dateTimeOffset.ToUnixTimeSeconds();
            long oddL1 = ts1 % 10;
            ts1 = ts1 - oddL1;
            string tokenScret = MailCommonBase.HMACSHA1Encode(ts1.ToString(), Scret);
            return tokenScret; 
        }

       public class MailJobServiceJOB : IJob
        {
            private ILogger<MailJobServiceJOB> logger1;
            public Task Execute(IJobExecutionContext context)
            {
                Random r = new Random();
                int randomInteger = r.Next(1, 9999);
                Thread.Sleep(randomInteger);

                var services = new ServiceCollection();
                MailJobServiceStartup.ConfigureServices(services); 
                // 创建ServiceProvider
                var serviceProvider = services.BuildServiceProvider();
                var logger = serviceProvider.GetRequiredService<ILogger<MailJobServiceJOB>>();
                logger1 = logger;

                var tokenManagement = serviceProvider.GetRequiredService<IOptions<TokenManagement>>().Value;
                string dynamicToken = CreateDynamicToken(tokenManagement.Secret);
                var mailListApiUrl = serviceProvider.GetRequiredService<IOptions<MailListApiUrl>>().Value;

                var globalConfigFromMainServiceAppSetting = serviceProvider.GetRequiredService<IOptions<GlobalConfigFromMainServiceAppSetting>>().Value;
                Console.Write("IshopX的Email服务的配置是来自IshopX网站,而不是本地的AppSetting.json或者EmailServiceConfig.json 本地而不是本地的AppSetting.json的GlobalConfig节点如下:\n");
                Console.Write(JsonConvert.SerializeObject(globalConfigFromMainServiceAppSetting,Formatting.Indented));
                Console.Write("\n------------------------------------------------------------------------------------------------\n");

                logger.LogInformation("[TokenManagement::Secret = {0} {1}:{2}] - [DynamicToken = {3}]", tokenManagement.Secret, mailListApiUrl.HostApiUrl, mailListApiUrl.Port, dynamicToken);

                logger.LogInformation("[TokenManagement::Secret = {0} {1}:{2}] - [DynamicToken = {3}]", tokenManagement.Secret, mailListApiUrl.HostApiUrl, mailListApiUrl.Port, dynamicToken);
                 
                MailTaskJobRequest mailTaskJobRequest =  GetMailTaskJob(mailListApiUrl, dynamicToken);
                //把 邮件对象获取下来发送 
                try
                {   
                    logger.LogInformation("[App is start running ]");
                    logger.LogError("[MailJobServiceJOB SENDING BEGIN]");
                     
                    if(mailTaskJobRequest.Success)
                    {
                        mailTaskJobRequest.EmailBody = MailCommonBase.Base64ToString(mailTaskJobRequest.EmailBody);
                        EmailJob emailJob = new EmailJob(mailTaskJobRequest.SendMailInfo);
                        emailJob.Run(mailTaskJobRequest).GetAwaiter();
                        logger.LogInformation("OK...........................................................................[SUCCESS={0}]", mailTaskJobRequest.Success);
                    }
                    else
                    {
                        logger.LogError("\n[{0}][MailTaskJobRequest.Success = FALSE] [RETURN ERROR:{1}]",DateTime.Now, mailTaskJobRequest.Subject);
                    }
                }catch(Exception ex)
                {
                    logger.LogError("[MailJobServiceJOB.Execute]", ex.Message);
                }
                 
                return Console.Out.WriteLineAsync(string.Format("\n[{0:yyyy-MM-dd HH:mm:ss fff}] [INFO] [Execute::MailJobServiceJOB:{1}]", DateTime.Now, context.FireInstanceId));
            }

            public void ExecuteTest()
            {
                Random r = new Random();
                int randomInteger = r.Next(1, 9999);
                Thread.Sleep(randomInteger);

                var services = new ServiceCollection();
                MailJobServiceStartup.ConfigureServices(services);
                // 创建ServiceProvider
                var serviceProvider = services.BuildServiceProvider();
                var logger = serviceProvider.GetRequiredService<ILogger<MailJobServiceJOB>>();
                 
                var tokenManagement = serviceProvider.GetRequiredService<IOptions<TokenManagement>>().Value;
                string dynamicToken = CreateDynamicToken(tokenManagement.Secret);
                var mailListApiUrl = serviceProvider.GetRequiredService<IOptions<MailListApiUrl>>().Value;

                logger.LogInformation("[TokenManagement::Secret = {0} {1}:{2}] - [DynamicToken = {3}]", tokenManagement.Secret, mailListApiUrl.HostApiUrl, mailListApiUrl.Port, dynamicToken);

                MailTaskJobRequest mailTaskJobRequest = GetMailTaskJob(mailListApiUrl, dynamicToken);
                //把 邮件对象获取下来发送 
                try
                {
                    logger.LogInformation("[App is start running ]");
                    logger.LogError("[MailJobServiceJOB SENDING BEGIN]");

                    if (mailTaskJobRequest.Success)
                    {
                        EmailJob emailJob = new EmailJob(mailTaskJobRequest.SendMailInfo);
                        emailJob.Run(mailTaskJobRequest).GetAwaiter();
                        logger.LogInformation("OK.............................................................................................");
                    }
                    else
                    {
                        logger.LogError("[MailTaskJobRequest.Success = FALSE]", DateTime.Now);
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError("[MailJobServiceJOB.Execute]", ex.Message);
                }
                //return Console.Out.WriteLineAsync(string.Format("\n[{0:yyyy-MM-dd HH:mm:ss fff}] [INFO] [Execute::MailJobServiceJOB:{1}]", DateTime.Now, context.FireInstanceId));
            }
            public MailTaskJobRequest GetMailTaskJob(MailListApiUrl mailListApiUrl, string dynamicToken)
            {
                mailListApiUrl.HostApiUrl = mailListApiUrl.HostApiUrl.TrimEnd('/');
                string  mailTaskUrl = string.Format("{0}:{1}/Mgr/ShopAdmin/GetMailTaskJobRequest/{2}", mailListApiUrl.HostApiUrl, mailListApiUrl.Port, dynamicToken);
                 
                using (HttpClient client = new HttpClient())
                {
                    string responseBody = "";
                    try
                    {
                        HttpResponseMessage response = client.GetAsync(mailTaskUrl).Result; 
                        response.EnsureSuccessStatusCode();
                        responseBody = response.Content.ReadAsStringAsync().Result;
                        MailTaskJobRequest mailTaskJobRequest = JsonConvert.DeserializeObject<MailTaskJobRequest>(responseBody);
                        return mailTaskJobRequest;
                    }
                    catch (Exception ex)
                    {
                        string errSocketEx = string.Format("[{0:yyyy-MM-dd HH:mm:ss}][MailJobServiceJOB.GetMailTaskJob()][NETWORK ERROR][Exception]",DateTime.Now, ex.Message);
                        MailCommonBase.OperateDateLoger(errSocketEx);
                        MailTaskJobRequest mailTaskJobRequestEx = new MailTaskJobRequest
                        {
                            Success = true,
                            SendMailInfo = null,
                            Subject = null,
                            ToMailArray = null,
                            EmailBody = null
                        };
                        logger1.LogError("[MailJobServiceJOB.GetMailTaskJob()][Exception][{0}]",ex.Message);
                        return mailTaskJobRequestEx;
                    }
                }
            }
        }

        public class MailJobServiceStartup
        {
            public static void ConfigureServices(IServiceCollection services)
            {
                // 创建 appsettings
                var Configuration = ReadFromAppSettings();

                // 配置日志
                services.AddLogging(builder =>
                {
                    builder.AddConsole();
                    builder.AddDebug();
                });
                
                services.Configure<TokenManagement>(Configuration.GetSection("TokenManagement")); 
                services.AddOptions<TokenManagement>("TokenManagement");
                services.AddOptions<TokenManagement>("globalConfig");
                services.Configure<MailListApiUrl>(Configuration.GetSection("MailListApiUrl"));
                services.AddOptions<MailListApiUrl>("MailListApiUrl");

                services.AddTransient<TokenManagement>();
            }
            public static IConfigurationRoot ReadFromAppSettings()
            {
                return new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", false)
                    .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("NETCORE_ENVIRONMENT")}.json", optional: true)
                    .AddEnvironmentVariables()
                    .Build();
            }
        }

        public class EmailJob
        {
            public EmailJob(SendMailInfo sendMailInfo)
            {
                SendMailInfo = sendMailInfo;
               
            }
            public SendMailInfo SendMailInfo { get; set; }
             
            public async Task Run(MailTaskJobRequest mailTaskJobRequest)
            {
                #region BEGIN
                
                foreach (var item in mailTaskJobRequest.ToMailArray)
                {
                    string toMail = item;
                    if (!MailCommonBase.IsValidEmail(toMail))
                    {
                        continue;
                    }
                    #region EmailHelper
                    try
                    { 
                        await Task.Delay(1);
                          
                        EmailHelper email = new EmailHelper(SendMailInfo, toMail, mailTaskJobRequest.Subject, mailTaskJobRequest.EmailBody);

                        //email.AddAttachments(attachPath);

                        switch (SendMailInfo.SenderOfCompany)
                        {
                            case "163":
                                email.SendMail163(); 
                                break;
                            case "126":
                                email.SendMail126();
                                break;
                            case "QQ":
                                email.SendMailQQ();
                                break;
                            case "GOOGLE":
                                email.SendMailGoogle();
                                break;
                            default:
                                email.SendMail();
                                break;
                        }
                        Console.WriteLine("[{0:yyyy-MM-dd HH:mm:ss}] Sending Email （{1}）......", DateTime.Now, toMail);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                        MailCommonBase.OperateLoger(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][EmailJob.Run][EXCEPTION][{1}]",DateTime.Now, ex.Message));
                    }
                    #endregion
                }
                #endregion END 
            }
             
             
        }
    }
}
 

