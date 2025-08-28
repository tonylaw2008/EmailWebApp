 
using MailEnhanceService;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Net.Http;
using Microsoft.Extensions.Logging;

#region BATCH ITEM BEGIN

string[] cmdArgs = Environment.GetCommandLineArgs(); // 取得命令行參數
foreach (var item in cmdArgs)
{
    Console.WriteLine($"cmdArgs Parameters: {item}\n");  
}

// 發送郵件列表
string[] mailToList;

if (args.Length == 0)
{
    Console.WriteLine("No any Parameter!");
    // 没有参数时，初始化长度为1的数组并赋值
    mailToList = new string[] { "mcessol2000@gmail.com" };
}
else
{
    // 根据参数数量初始化数组
    mailToList = new string[args.Length];
    int i = 0;
    foreach (var arg in args)  //args : 用戶傳入的參數
    {
        Console.WriteLine($"Parameter: {arg}");
        mailToList[i] = arg; // 直接使用索引赋值，更简洁
        i++;
    }
}

// 附件
string[] attachmentFileList = new string[] { @"D:\APP\DGX_NET6\STAR_NEW\EmailWebApp\README_IMGs\README\DMZ_SETTING.jpg",
                                             @"D:\APP\DGX_NET6\STAR_NEW\EmailWebApp\README_IMGs\README\EMAIL_SENDER_SETTTING.jpg" }; // 多个附件

// 作为控制台应用运行
Console.WriteLine("\nThe Console application is running now....\n\n");

Console.WriteLine("$ Command Mode: MailEnhanceService \"abc@gmail.com\"  \"abc@yahoo.com\" \n\nPress [ENTER] to continue......\n");

Console.ReadLine();

#if RELEASE
Console.WriteLine("\nIf you use the [Release] mode to test, please note: you need to copy the SSL certificate and email template to the corresponding release program.\n\n");
#endif

//方式 II
string? mainComId = string.Empty; // 非多公司平台，則可以不指定 mainComId
EmailAppService emailAppService = EmailAppService.StartUpEmailAppService(mainComId);
emailAppService._logger.LogInformation("EmailAppService has been started successfully.");
  
bool success = false;

success = await emailAppService.RunAsync(
    mailToList,
    "測試郵件主題", //如果沒有主題，則會使用 內容的純文本前20字作為主題。
    MailTemplateEnum.REGISTER,  //如果指定了郵件模版，則會使用指定的模板内容，而不會例會傳入參數 bodyRawContent 的內容
    "", // 不使用模板，必須 MailTemplateEnum.NO_TEMPLATE
    "en-US",//如果是使用模版，則使用什麼語言版本的模板。
    "http://192.168.0.9:8080/zh-HK/Device/CardDocBuild",
    attachmentFileList);


Console.WriteLine($"[測試郵件主題]-多個郵件地址發送整體結果: {success}");

  
#endregion BATCH ITEM END
 
Console.ReadLine();

