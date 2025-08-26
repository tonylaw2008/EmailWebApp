using log4net;
using Microsoft.AspNetCore.Builder;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace EmailWebApp
{
    public static class LoggerServiceExt
    {
        /// 使用log4net配置
        /// <param name="app"></param>
        /// <returns></returns>
        public static IApplicationBuilder UseLog4net(this IApplicationBuilder app)
        {
            var logRepository = log4net.LogManager.CreateRepository(AppDomain.CurrentDomain.BaseDirectory, typeof(log4net.Repository.Hierarchy.Hierarchy));
            log4net.Config.XmlConfigurator.Configure(logRepository, new FileInfo("log4net.config"));
            return app;
        }
    }
    public static class Log4NetHelper
    {
        //log4net日志初始化 这里就读取多个配置节点

        private static readonly ILog _logError = LogManager.GetLogger(Assembly.GetCallingAssembly(), "LogError");
        private static readonly ILog _logNormal = LogManager.GetLogger(Assembly.GetCallingAssembly(), "LogNormal");
        private static readonly ILog _logAOP = LogManager.GetLogger(Assembly.GetCallingAssembly(), "LogAOP");

        /// <summary>
        /// 日誌記錄錯誤資訊,帶異常
        /// </summary>
        /// <param name="info"></param>
        /// <param name="ex"></param>
        public static void Log4NetException(string info, Exception ex)
        {
            if (_logError.IsErrorEnabled)
            {
                Console.WriteLine(info,ex);
                _logError.Error(info, ex);
            }
        }
        /// <summary>
        /// 日誌記錄錯誤資訊
        /// </summary>
        /// <param name="message"></param>
        public static void Log4NetError(string message)
        {
            if (_logError.IsErrorEnabled)
            {
                Console.WriteLine(message);
                _logError.Error(message);
            }
        }

        /// <summary>
        /// 日誌記錄INFO資訊
        /// </summary>
        /// <param name="info"></param>
        public static void Log4NetInfo(string info)
        {
            if (_logNormal.IsInfoEnabled)
            {
                Console.WriteLine(info);
                _logNormal.Info(info);
            }
        }

        /// <summary>
        /// 日誌記錄AOP資訊,帶key
        /// </summary>
        /// <param name="key"></param>
        /// <param name="info"></param>
        public static void Log4NetAOP(string key, string info)
        {
            if (_logAOP.IsInfoEnabled)
            {
                Console.WriteLine($"{key}:{info}");
                _logAOP.Info($"{key}:{info}");
            }
        }
        /// <summary>
        /// 日誌记录AOP信息,不带key
        /// </summary>
        /// <param name="info"></param>
        public static void Log4NetAOP(string info)
        {
            if (_logAOP.IsInfoEnabled)
            {
                Console.WriteLine(info);
                _logAOP.Info(info);
            }
        }
    }
}
