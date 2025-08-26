using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO; 
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection; 
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Threading.Tasks; 
using System.Text;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Hosting; 
using Microsoft.Extensions.Configuration.Json;
using Microsoft.Extensions.Options;
using System.Net;
using System.Threading;
using static MailJobService.MailJobServiceQuartz;

namespace MailJobService
{
    
    class Program
    { 
        static void Main(string[] args)
        { 
            //MailJobServiceJOB mailJobServiceJOB = new MailJobServiceJOB(); //test
            //mailJobServiceJOB.ExecuteTest();                               //test
            //==============================================================================================================================
            Console.Clear();
            Console.OutputEncoding = Encoding.UTF8;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
 
            if (args.Length == 0)
            {
                args = new string[1];
                args.SetValue("caihaili82@gmail.com", 0); 
            }
            bool IsValidEmail = MailCommonBase.IsValidEmail(args[0]);
            Console.WriteLine("[IsValidEmail] {0} = {1}", args[0], IsValidEmail);
             
            //EmailApp SETTING FROM [EmailServiceConfig.json]
            //BEGIN===============================================================================================================
            //BEGIN===============================================================================================================
            Console.WriteLine("\n[START [EmailApp EmailServiceConfig.json] TEST, Press Key Y ]\n");
            ConsoleKey ConsoleKeyInputStart = Console.ReadKey().Key;
            if(ConsoleKeyInputStart == ConsoleKey.Y)
            {
                Console.WriteLine("\nMailService::Hello World! EmailServiceConfig.json\n");

                // EmailServiceConfig.json 【EmailApp】是使用独立配置文件版本
                if (!File.Exists(Path.Combine(Directory.GetCurrentDirectory(), "EmailServiceConfig.json")))
                {
                    Console.WriteLine("\nFILE NOT EXIST [EmailServiceConfig.json]");
                    MailCommonBase.OperateLoger("[Program.Main] FILE NOT EXIST [EmailServiceConfig.json]");
                }
                else
                {
                    string EmailServiceConfigJsonInUsed = "\n[EmailServiceConfig.Json InUse For test send Email]";
                    Console.WriteLine("{0} {1}", DateTime.Now, EmailServiceConfigJsonInUsed);
                }

                var EmailServiceConfigContent = MailCommonBase.ReadDataFromJson(Directory.GetCurrentDirectory(), "EmailServiceConfig.json");
                EmailServiceConfig emailServiceConfig = JsonConvert.DeserializeObject<EmailServiceConfig>(EmailServiceConfigContent);

                EmailApp emailApp = new EmailApp(emailServiceConfig.SendMailInfo, emailServiceConfig.GmailInfo);

                try
                {
                    string LanguageCode = "zh-CN";
                    string subject = $"ACCOUNT REGISTER EMAIL COMFIRMED";
                    string attachPath = "";

                    emailApp.Run(args, subject, MailTemplateEnum.Register.ToString(), LanguageCode, "https://www.microsoft.com", attachPath).GetAwaiter();

                    Console.WriteLine("\nOK,Push to asyc task............................................................................................\n");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("{0} {1}", ex.Message, "[APP RUNNING EXCEPTION ERROR OCCURED ]\n");
                    MailCommonBase.OperateDateLoger("[APP RUNNING EXCEPTION ERROR OCCURED ]" + ex.Message);
                }
            }
            Thread.Sleep(3000);
            //Service Inject Test
            //BEGIN===============================================================================================================
            //BEGIN===============================================================================================================
            Console.WriteLine("\n[START [App.cs] Service Inject Test, Press Key X ]\n");
            ConsoleKey ConsoleKeyServiceInjectStart = Console.ReadKey().Key;
            if (ConsoleKeyServiceInjectStart == ConsoleKey.X)
            {
                // 创建ServiceCollection
                var services = new ServiceCollection();
                ConfigureServices(services);

                // 创建ServiceProvider
                var serviceProvider = services.BuildServiceProvider();
                var logger = serviceProvider.GetRequiredService<ILogger<Program>>();

                //var sendMailInfo = serviceProvider.GetRequiredService<IOptions<SendMailInfo>>();
                //logger.LogInformation("Program::sendMailInfo = ", sendMailInfo);

                try
                {
                    string LanguageCode = "zh-HK";
                    Console.WriteLine("\n[INPUT LANGUAGE CODE (zh-CN, zh-HK, en-US)]:\n");
                    LanguageCode = Console.ReadLine();
                    LanguageCode??= "zh-HK";

                    string Subject = $"ACCOUNT REGISTER EMAIL COMFIRMED";
                    Console.WriteLine("\n[INPUT SUBJECT (e.g.: ACCOUNT REGISTER EMAIL COMFIRMED),If Leave blank it will be from Email Template]:\n");
                    Subject = Console.ReadLine();

                    string attachPath = "";
                    logger.LogInformation("\n[App is start running ]");

                    Task appTask = serviceProvider.GetRequiredService<App>().Run(args, Subject, MailTemplateEnum.Register, LanguageCode, "https://www.microsoft.com", attachPath);  //use GetRequiredService,not GetService(),if not registed,it will thorw exception. 

                    Console.WriteLine("OK.............................................................................................");  
                }
                catch (Exception ex)
                {
                    logger.LogInformation(ex, "[APP RUNNING EXCEPTION ERROR OCCURED ]");
                    MailCommonBase.OperateDateLoger("[APP RUNNING EXCEPTION ERROR OCCURED ]" + ex.Message);
                }
            }
            Thread.Sleep(3000);
            //JOB SCHEDULE
            //BEGIN===============================================================================================================
            //BEGIN===============================================================================================================
            //API 获取发送邮件列表与配置后本地负责发送服务
            //MAIL SERVICE BEGIN
            DateTimeOffset dateTimeOffset = new DateTimeOffset(DateTime.Now);
            long ts1 = dateTimeOffset.ToUnixTimeSeconds(); 
            long oddL1 = ts1 % 10;
            long ts2 = ts1 - oddL1; 
            Console.WriteLine("\n[MAIL SERVICE BEGIN.......][{0}] [{1}] [{2}]\n", ts1, ts2, oddL1);

            var configuration = ReadFromAppSettings();
            int IntervalMinutes = int.Parse(configuration.GetSection("MailListApiUrl:IntervalMinutes").Value);
            MailJobServiceQuartz.RunScheduleMonthlyGlobalProgram(DateTime.Now, IntervalMinutes).GetAwaiter();

            Console.ReadKey();
            Console.ReadLine();
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

            var sendMailInfo = Configuration.GetSection("SendMailInfo");
            services.Configure<SendMailInfo>(sendMailInfo);

            var gmailInfo = Configuration.GetSection("GmailInfo");
            services.Configure<GmailInfo>(gmailInfo);

            services.AddOptions<SendMailInfo>("SendMailInfo");
            services.AddOptions<SendMailInfo>("GmailInfo");

            services.AddTransient<SendMailInfo>();
            services.AddTransient<GmailInfo>();
            services.AddTransient<App>();

            // 配置日志
            services.AddLogging(builder =>
            {
                builder.AddConsole();
                builder.AddDebug();
            });
        }
         
    }
     
}
