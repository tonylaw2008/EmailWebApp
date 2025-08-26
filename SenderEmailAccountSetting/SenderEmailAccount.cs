using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SenderEmailAccountSetting
{
    public class SenderEmailAccount
    {
        /// <summary>
        /// 發件人公司名稱（例如：yahoo.hk, 163.com, 126.com, qq.com, gmail.com等）
        /// 域名部分
        /// </summary>
        public string SenderOfCompany { get; set; }

        /// <summary>
        /// 使用那個工具發送EMAIL:  0: System.Net.Mail 系統自帶, 1: MailKit
        /// </summary>
        public int MailTool { get; set; }


        /// <summary>
        /// 是否啟用 SSL（安全套接層）協議   
        /// </summary>
        public bool EnableSSL { get; set; }

        /// <summary>
        /// 是否啟用 TSL（傳輸層安全性）協議
        /// </summary>
        public bool EnableTSL { get; set; }
        public bool EnablePasswordAuthentication { get; set; }

        /// <summary>
        /// SMTP 伺服器主機地址
        /// </summary>
        public string SenderServerHost { get; set; }

        /// <summary>
        /// SMTP 伺服器主機端口號
        /// </summary>
        public int SenderServerHostPort { get; set; }

        /// <summary>
        /// 發件人郵箱地址
        /// </summary>
        public string FromMailAddress { get; set; }

        /// <summary>
        /// 發件人郵箱顯示名稱
        /// </summary>
        public string FromMailDisplayName { get; set; }

        /// <summary>
        /// SMTP 認證登錄賬號
        /// </summary>
        public string SenderUserName { get; set; }

        /// <summary>
        /// SMTP 認證登錄密碼
        /// </summary>
        public string SenderUserPassword { get; set; }

        /// <summary>
        /// 備註信息
        /// </summary>
        public string Remarks { get; set; }
    }
}
