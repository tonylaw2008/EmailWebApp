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

namespace MailJobService
{
    
    class ProgramXX
    { 
        static void MainXX(string[] args)
        {
            //==============================================================================================================================
            Console.Clear();
            Console.OutputEncoding = Encoding.UTF8;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); 
#if DEBUG
            if (args.Length == 0)
            {
                args = new string[2];
                args.SetValue("LawtatfaiTony@gmail.com", 0);
                args.SetValue("LawtatfaiTony@qq.com", 1);
            }
            bool IsValidEmail = MailCommonBase.IsValidEmail(args[0]);
            Console.WriteLine("[IsValidEmail] {0} = {1}", args[0], IsValidEmail);

            bool IsValidEmail1 = MailCommonBase.IsValidEmail(args[1]);
            Console.WriteLine("[IsValidEmail] {0} = {1}", args[1], IsValidEmail1);
#endif
            Console.WriteLine("IdentityUtility::Hello World! EmailServiceConfig.json");

            // EmailServiceConfig.json 【EmailApp】是使用独立配置文件版本
            var EmailServiceConfigContent = MailCommonBase.ReadDataFromJson(Directory.GetCurrentDirectory(), "EmailServiceConfig.json");
            EmailServiceConfig emailServiceConfig = JsonConvert.DeserializeObject<EmailServiceConfig>(EmailServiceConfigContent);

            EmailApp emailApp = new EmailApp(emailServiceConfig.SendMailInfo, emailServiceConfig.GmailInfo); 

            try
            {
                //serviceProvider.GetService<App>().Run(args);
                string LanguageCode = "zh-CN";
                string subject = $"ACCOUNT REGISTER EMAIL COMFIRMED";
                Console.WriteLine("\nInput Email Subject:\n");
                Console.WriteLine("If use the begin 20 words of Html template text,then input Y:\n");
                subject = Console.ReadLine();
                
                bool isQuit = false;
                while (isQuit == false)
                {
                    if (!string.IsNullOrEmpty(subject))
                    {
                        isQuit = true;
                    }
                    else
                    {
                        Console.WriteLine("\nInput Email Subject:\n");
                        Console.WriteLine("If use the begin 20 words of Html template text,then input Y:\n");
                        subject = Console.ReadLine();
                        isQuit = false;
                    }
                }
                string attachPath = "";

                emailApp.Run(args, subject, MailTemplateEnum.Register.ToString(), LanguageCode, "https://www.microsoft.com", attachPath).GetAwaiter();
                
                Console.WriteLine("OK,Pushed to asyc task............................................................................................");
                Console.ReadKey(); 
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0} {1}", ex.Message, "[APP RUNNING EXCEPTION ERROR OCCURED ]");
                MailCommonBase.OperateDateLoger("[APP RUNNING EXCEPTION ERROR OCCURED ]" + ex.Message);
            }
        }
    }
}
