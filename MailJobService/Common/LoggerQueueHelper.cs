using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using static MailJobService.MailCommonBase;

namespace MailJobService
{
    public class Logger
    {
        public LoggerMode loggerMode { get; set; } = LoggerMode.INFO;
        public string LoggerLine { get; set; } = string.Empty;
    }
    public class LoggerQueueHelper
    {
        private static string _LogPath;
        public static string LogPath
        {
            get
            {
                if (MailMemoryCacheHelper.Contains("LogPath") == false)
                {
                    DateTimeOffset thisOffsetTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,23,59,59);
                    if (IsLowVsersion())
                    {
                        MailMemoryCacheHelper.Set("LogPath", BasePath, thisOffsetTime);

                        string path1 = Path.Combine(BasePath, "log");
                        if (!Directory.Exists(path1))
                        {
                            Directory.CreateDirectory(path1);
                        }
                        _LogPath = BasePath;
                    }
                    else
                    {
                        _LogPath = AppDomain.CurrentDomain.BaseDirectory;
                        MailMemoryCacheHelper.Set("LogPath", _LogPath, thisOffsetTime);
                    }
                }
                else
                {
                    _LogPath = MailMemoryCacheHelper.GetCacheItem<string>("LogPath");
                }
                return _LogPath;
            }
        }
        #region  
        private static object lockObj1 = new object();
        private static void Log(Logger logger)
        {
            lock (lockObj1)
            {
                string filePath = Path.Combine(LogPath, "log");

                string filename = string.Format("LoggerInfo{0:yyyyMMdd}.log", DateTime.Now);
                string file = Path.Combine(filePath, filename);
                switch (logger.loggerMode)
                {
                    case LoggerMode.DEBUG:
                        filename = string.Format("LoggerDebug{0:yyyyMMdd}.log", DateTime.Now);
                        file = Path.Combine(filePath, filename);
                        break;
                    case LoggerMode.ERROR:
                        filename = string.Format("LoggerError{0:yyyyMMdd}.log", DateTime.Now);
                        file = Path.Combine(filePath, filename);
                        break;
                    case LoggerMode.FATAL:
                        filename = string.Format("LoggerFatal{0:yyyyMMdd}.log", DateTime.Now);
                        file = Path.Combine(filePath, filename);
                        break;
                    case LoggerMode.WARNNING:
                        filename = string.Format("LoggerWarnning{0:yyyyMMdd}.log", DateTime.Now);
                        file = Path.Combine(filePath, filename);
                        break;
                    default:
                        break;
                }
                using (StreamWriter w = new StreamWriter(file, true, Encoding.UTF8))
                {
                    w.WriteLine(logger.LoggerLine);
                    w.Close();
                }
            }
        }
        #endregion
        private static object lockObj2 = new object();
        private static void ErrorLog(string msg)
        {
            lock (lockObj2)
            {
                string filePath = Path.Combine(LogPath, "log");

                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }
                string file = Path.Combine(filePath, "LogError.log");

                using (StreamWriter w = new StreamWriter(file, true, Encoding.UTF8))
                {
                    w.WriteLine(msg);
                    w.Close();
                }
            }
        }

        /// <summary>
        ///  Low version Basic on Win10
        /// </summary>
        /// <returns></returns>
        public static bool IsLowVsersion()
        {
            Version currentVersion = Environment.OSVersion.Version;
            Version win10Version = new Version("6.2");
            if (currentVersion < win10Version)
                return true;
            else
                return false;
        }

        public bool AnalysisStatus = false;

        public static LoggerQueueHelper instance = new LoggerQueueHelper();
        private LoggerQueueHelper()
        {
        }

        private ConcurrentQueue<Logger> t = new ConcurrentQueue<Logger>();
        /// <summary>
        /// add to queue
        /// </summary>
        /// <param name="tip"></param>
        public int AddData(Logger tip)
        {
            t.Enqueue(tip);
            return t.Count;
        }

        /// <summary>
        /// get the the first queue
        /// </summary>
        /// <returns></returns>
        public Logger Get()
        {
            Logger logger = new Logger();
            return t.TryPeek(out logger) ? logger : null;
        }

        /// <summary>
        ///get from queue and calc
        /// </summary>
        public void PostAnalysis()
        {
            while (t.Count > 0)
            {
                try
                {
                    AnalysisStatus = true;

                    Logger outresult = new Logger();

                    if (t.TryDequeue(out outresult))
                    {
                        Log(outresult);
                    }
                    else
                    {
                        ErrorLog(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][ERROR (TryDequeue error)]", DateTime.Now));
                    }
                }
                catch (Exception ex)
                {
                    AnalysisStatus = false;
                    ErrorLog(string.Format("[{0:yyyy-MM-dd HH:mm:ss}][EXCEPTION] [{1}]", DateTime.Now, ex.Message));
                    throw;
                }
            }
            AnalysisStatus = false;
        }

        /// <summary>
        /// get the queue total count
        /// </summary>
        /// <returns></returns>
        public int Getcount()
        {
            return t.Count;
        }
    }
}