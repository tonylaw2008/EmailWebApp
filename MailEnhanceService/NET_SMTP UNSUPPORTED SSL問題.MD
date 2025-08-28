# System.Net.Mail.SmtpClient 嗯支持SSL



https://ithelp.ithome.com.tw/articles/10369379



為了替各個系統建立測試 EMail 發信的環境，我使用 hMailServer 建立收發 EMail 伺服器，並用 openssl 建立自簽憑證供 SSL 使用。

使用 .NET 8 的 SmtpClient 測試發信，測 25 port 發信不啟用 SSL （明文通道發信）－理所當然地成功，但測發信到 465 port + 啟用 SSL/TLS 就是一直 Timeout，找遍了 EMail 伺服器設置、SSL 證書及 TLS 設定都找不到任何原因...又噴了一個下午。

接近下班時間，趕快往下測 587 port + StartTLS 發信並啟用SSL，結果一下就成功，證明下午的查各種設置、設定都搞錯方向了，問題就是「EMail 伺服器啟用 SSL/TLS」！

google「.net smtp client tls」果然在[SMTP and Implicit TLS from a C# .NET application](https://medium.com/@adrian.nasui/smtp-and-implicit-tls-from-a-c-net-application-c451609b7c46) 找到了正解！

結論就是 .NET SmtpClient：

1. 支援沒加密的明文傳輸。
2. 不支援 SSL/TLS 連接安全性（所謂的 implicit TLS)，不管是常用的 465 port 或其他自定義 port，都只會連線 Timeout。
3. 支援 StartTLS 連接安全性（所謂的 explicit TLS)，測試 25 port 和常用的 587 port 皆可。
4. 如果必須用SSL/TLS 連接安全性（所謂的 implicit TLS)，可改用第三方元件例如：[MailKit](https://www.nuget.org/packages/MailKit/)。



★★★  https://stackoverflow.com/questions/26414424/which-versions-of-ssl-tls-does-system-net-webrequest-support

其他

When using System.Net.WebRequest your application will negotiate with the server to determine the highest TLS version that both your application and the server support, and use this. You can see more details on how this works here:

http://en.wikipedia.org/wiki/Transport_Layer_Security#TLS_handshake

If the server doesn't support TLS it will fallback to SSL, therefore it could potentially fallback to SSL3. You can see all of the versions that .NET 4.5 supports here:

http://msdn.microsoft.com/en-us/library/system.security.authentication.sslprotocols(v=vs.110).aspx

In order to prevent your application being vulnerable to POODLE, you can disable SSL3 on the machine that your application is running on by following this explanation:

https://serverfault.com/questions/637207/on-iis-how-do-i-patch-the-ssl-3-0-poodle-vulnerability-cve-2014-3566