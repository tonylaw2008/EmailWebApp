using System;
using System.Collections.Generic;
using System.Text;

namespace MailJobService
{
    public class AppSettingModel
    {
       public ConnectionConfig connectionConfig;
    }
    public class ConnectionConfig
    { 
        private string _Server;
        public string Server
        {
            get
            {
                return _Server;
            }
            set
            {
                _Server = value;
            }
        }
        private string _DataBase;
        public string DataBase
        {
            get
            {
                return _DataBase;
            }
            set
            {
                _DataBase = value;
            }
        }
        private bool _IsTrustMode;
        public bool IsTrustMode
        {
            get
            {
                return _IsTrustMode;
            }
            set
            {
                _IsTrustMode = value;
            }
        }

        private string _UserId;
        public string UserId
        {
            get
            {
                return _UserId;
            }
            set
            {
                _UserId = value;
            }
        }
        private string _Password;
        public string Password
        {
            get
            {
                return _Password;
            }
            set
            {
                _Password = value;
            }
        }

        private string _ConnectiongString;
        public virtual string ConnectionString
        {
            get
            {
                ////---------------------------------------------------------------------------------------------------
                if (_IsTrustMode)
                {
                    _ConnectiongString = string.Format("Server = {0}; Database = {1}; Trusted_Connection = True;Integrated Security=SSPI", Server, DataBase);
                }
                else
                {
                    _ConnectiongString = string.Format("Data Source={0};Initial Catalog={1};User ID={2};Password={3};TrustServerCertificate=True;ApplicationIntent=ReadWrite;", Server, DataBase, UserId, Password);
                }
                return _ConnectiongString;
            }
        } 
    }
}
