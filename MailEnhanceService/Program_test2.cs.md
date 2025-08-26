 
using MailEnhanceService;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Net.Http;
using Microsoft.Extensions.Logging;

// 作为控制台应用运行
Console.WriteLine("The Console application is running now....");

//方式 I 
//// 创建服务集合
//var services = new ServiceCollection();
//ServiceConfigurator.ConfigureServices(services);

//// 构建服务提供者
//var serviceProvider = services.BuildServiceProvider();

//// 获取邮件服务
//using var scope = serviceProvider.CreateScope();
//EmailAppService emailAppService = scope.ServiceProvider.GetRequiredService<EmailAppService>();

//方式 II
string? mainComId = string.Empty; // 非多公司平台，則可以不指定 mainComId
EmailAppService emailAppService = EmailAppService.StartUpEmailAppService(mainComId);
emailAppService._logger.LogInformation("EmailAppService has been started successfully.");

#region BATCH ITEM BEGIN

// 发送测试邮件
string[] mailToList = { "13392226@qq.com" };
bool success = false;

success = await emailAppService.RunAsync(
    mailToList,
    "測試郵件主題", //如果沒有主題，則會使用 內容的純文本前20字作為主題。
    MailTemplateEnum.REGISTER,  //如果指定了郵件模版，則會使用指定的模板内容，而不會例會傳入參數 bodyRawContent 的內容
    "", // 不使用模板，必須 MailTemplateEnum.NO_TEMPLATE
    "en-US",//如果是使用模版，則使用什麼語言版本的模板。
    "http://192.168.0.9:8080/zh-HK/Device/CardDocBuild",
    null);

Console.WriteLine($"[測試郵件主題]-郵件發送結果: {success}");

//IList<SenderEmailAccount> SenderEmailAccountList = emailAppService.GetSenderEmailAccountList();
//SenderEmailAccount SenderEmailAccount = SenderEmailAccountList[0];

//using SmtpClient client = new SmtpClient();

//client.Host = SenderEmailAccount.SenderServerHost;

//client.Port = SenderEmailAccount.SenderServerHostPort;
//client.EnableSsl = true; // 启用 SSL
//client.Timeout = 30000; // 设置超时
//client.UseDefaultCredentials = false;


//client.Credentials = new NetworkCredential(
//    SenderEmailAccount.SenderUserName,
//    SenderEmailAccount.SenderUserPassword
//);

//MailMessage mailMessage = new MailMessage();
//mailMessage.From = new MailAddress(SenderEmailAccount.FromMailAddress);
//mailMessage.Subject = "Email Test Subject";
//mailMessage.Body = "Email Test Subject";
//mailMessage.To.Add("caihaili82@gmail.com");
//mailMessage.IsBodyHtml = true;
//mailMessage.SubjectEncoding = Encoding.UTF8;
//mailMessage.BodyEncoding = Encoding.UTF8;
//mailMessage.Priority = MailPriority.High;
//mailMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnSuccess;
//mailMessage.Sender = new MailAddress(SenderEmailAccount.FromMailAddress);

//client.Send(mailMessage);
//success = true;
//Console.WriteLine($"TESTII:郵件發送結果: {success}");

#endregion BATCH ITEM END



Console.ReadLine();

