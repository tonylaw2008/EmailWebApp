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

namespace MailJobService
{
    
    class Program
    { 
        static void Main(string[] args)
        { 
#if DEBUG
            if (args.Length == 0)
            {
                args = new string[2];
                args.SetValue("caihaili82@gmail.com", 0);
                args.SetValue("LawtatfaiTony@qq.com", 1);
            }
            bool IsValidEmail = CommonBase.IsValidEmail(args[0]);
            Console.WriteLine("{0} = {1}", args[0], IsValidEmail);

            bool IsValidEmail1 = CommonBase.IsValidEmail(args[1]);
            Console.WriteLine("{0} = {1}", args[1], IsValidEmail1);
            Console.WriteLine("IdentityUtility::Hello World!");
#endif

            //BEGIN===============================================================================================================
            //BEGIN===============================================================================================================
            // 创建ServiceCollection
            var services = new ServiceCollection();
            ConfigureServices(services);
             
            // 创建ServiceProvider
            var serviceProvider = services.BuildServiceProvider();
            var logger = serviceProvider.GetRequiredService<ILogger<Program>>();

            var sendMailInfo = serviceProvider.GetRequiredService<IOptions<SendMailInfo>>();
            CommonBase.OperateDateLoger("Program::sendMailInfo = ", sendMailInfo);
             
            try
            {
                //serviceProvider.GetService<App>().Run(args);
                string LanguageCode = "zh-HK";
                string subject = $"ACCOUNT REGISTER EMAIL COMFIRMED";
                string attachPath = "";
                logger.LogInformation("[App is start running ]");
                 
                Task appTask = serviceProvider.GetRequiredService<App>().Run(args, subject, MailTemplateEnum.Register, LanguageCode, "https://www.microsoft.com", attachPath);  //use GetRequiredService,not GetService(),if not registed,it will thorw exception. 

                Console.WriteLine("OK.............................................................................................");
                Console.ReadKey();

            }
            catch (Exception ex)
            {
                logger.LogInformation(ex, "[APP RUNNING EXCEPTION ERROR OCCURED ]");
                CommonBase.OperateDateLoger("[APP RUNNING EXCEPTION ERROR OCCURED ]" + ex.Message);
            }
        }
        public static IConfigurationRoot ReadFromAppSettings()
        {  
            return new ConfigurationBuilder()
                         .SetBasePath(Directory.GetCurrentDirectory())
                         .AddJsonFile("appsettings.json",false)
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
