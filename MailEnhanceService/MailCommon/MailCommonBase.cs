using System; 
using System.Collections.Generic; 
using System.Drawing; 
using System.IO;
using System.Linq;
using System.Net;
using System.Text;   
using System.Configuration; 
using System.Security.Cryptography;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Reflection;
using System.Runtime.InteropServices;

namespace MailEnhanceService
{
 
    public partial class MailCommonBase
    {
        private static string _BasePath;
        public static string BasePath
        {
            get
            { 
                _BasePath = AppDomain.CurrentDomain.BaseDirectory;
                if (MailLoggerQueueHelper.IsLowVsersion())
                {
                    _BasePath = Path.GetFullPath(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);
                    _BasePath = Path.GetDirectoryName(_BasePath);
                    return _BasePath;
                } 
                return _BasePath; 
            }
            set
            {
                _BasePath = value;
            }
        }
        public static string PathRemoveBin(string pathApp)
        {
            int pathIndex = pathApp.LastIndexOf("\\");
            if (pathIndex != -1)
            {
                string existBinPath = pathApp.Remove(0, pathIndex).ToLower();
                existBinPath = existBinPath.TrimStart('\\');

                if (existBinPath.ToLower() == "bin")
                {
                    pathApp = Directory.GetParent(pathApp).FullName;
                    return pathApp;
                }
                else
                {
                    return pathApp;
                }
            }
            else
            {
                return pathApp;
            }
        }
        public static string SHA1encode(string txt)
        {
            byte[] bytes = Encoding.ASCII.GetBytes(txt);
            HashAlgorithm hash = HashAlgorithm.Create(); 
            MemoryStream stream = new MemoryStream(bytes);
            return hash.ComputeHash(stream).Aggregate("", (s, e) => s + String.Format("{0:x2}", e), s => s);
        }
        public static string HMACSHA1Encode(string input, string strkey)
        {
            byte[] keyX = Encoding.ASCII.GetBytes(strkey);
            HMACSHA1 myhmacsha1 = new HMACSHA1(keyX);
            byte[] byteArray = Encoding.ASCII.GetBytes(input);
            MemoryStream stream = new MemoryStream(byteArray);
            return myhmacsha1.ComputeHash(stream).Aggregate("", (s, e) => s + String.Format("{0:x2}", e), s => s);
        }
        public static string HMACSHA1Encode(string input)
        {
            string strkey = DateTime.Now.Year.ToString();
            byte[] keyX = Encoding.ASCII.GetBytes(strkey);
            HMACSHA1 myhmacsha1 = new HMACSHA1(keyX);
            byte[] byteArray = Encoding.ASCII.GetBytes(input);
            MemoryStream stream = new MemoryStream(byteArray);
            return myhmacsha1.ComputeHash(stream).Aggregate("", (s, e) => s + String.Format("{0:x2}", e), s => s);
        }


        /// <summary>
        /// md5加密
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string md5Encode(string str)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] data = Encoding.Default.GetBytes(str);
            byte[] md5data = md5.ComputeHash(data);
            md5.Clear();
            string re_str = "";
            for (int i = 0; i < md5data.Length; i++)
            {
                re_str += md5data[i].ToString("x").PadLeft(2, '0');
            }
            return re_str;

        }
        public static string DoPost(string url, string data)
        {
            HttpWebRequest req = GetWebRequest(url, "POST");
            byte[] postData = Encoding.UTF8.GetBytes(data);
            Stream reqStream = req.GetRequestStream();
            reqStream.Write(postData, 0, postData.Length); reqStream.Close();
            HttpWebResponse rsp = (HttpWebResponse)req.GetResponse();
            Encoding encoding = Encoding.GetEncoding(rsp.CharacterSet);
            return GetResponseAsString(rsp, encoding);
        }
        public static HttpWebRequest GetWebRequest(string url, string method)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.ServicePoint.Expect100Continue = false;
            req.ContentType = "application/x-www-form-urlencoded;charset=utf-8";
            req.ContentType = "text/json";
            req.Method = method;
            req.KeepAlive = true;
            req.Timeout = 1000000;
            req.Proxy = null;
            return req;
        }
        public static string GetResponseAsString(HttpWebResponse rsp, Encoding encoding)
        {
            StringBuilder result = new StringBuilder();
            Stream stream = null;
            StreamReader reader = null;
            try
            {
                // 以字符流的方式读取HTTP响应
                stream = rsp.GetResponseStream();
                reader = new StreamReader(stream, encoding);
                // 每次读取不大于256个字符，并写入字符串
                char[] buffer = new char[256];
                int readBytes = 0;
                while ((readBytes = reader.Read(buffer, 0, buffer.Length)) > 0)
                {
                    result.Append(buffer, 0, readBytes);
                }
            }
            finally
            {
                if (reader != null) reader.Close();
                if (stream != null) stream.Close();
                if (rsp != null) rsp.Close();
            }
            return result.ToString();
        }
        public static DateTime ConvertIntDateTime(long unixDateTime)
        {

            DateTime start = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime date = start.AddMilliseconds(unixDateTime).ToLocalTime();
            return date;
        }
        public static DateTimeRangeObj DateTimeRangeParse(string strDateTimeRange)
        { 
            DateTimeRangeObj dateTimeRangeObj = new DateTimeRangeObj(); 
            string[] arrDateTimeRange = strDateTimeRange.Split(new char[] {'-', '/', 'T', ' ', '+'});
            List<string> list =new  List<string>();
            foreach(string item in arrDateTimeRange.ToList())
            { 
                if (item.Trim().Length > 0)
                    list.Add(item);
            }
            arrDateTimeRange = list.ToArray();

            string strStart,strEnd; 
            if (arrDateTimeRange.Count()==6)
            {
                strStart = string.Format("{0}-{1}-{2}", arrDateTimeRange[0], arrDateTimeRange[1], arrDateTimeRange[2]);
                if (!DateTime.TryParseExact(strStart, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTimeRangeObj.Start))
                {
                    dateTimeRangeObj.Start = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                }
                strEnd = string.Format("{0}-{1}-{2}", arrDateTimeRange[3], arrDateTimeRange[4], arrDateTimeRange[5]);
                if (!DateTime.TryParseExact(strEnd, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTimeRangeObj.End))
                {
                    dateTimeRangeObj.End = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                }
            }else
            {
                strStart = string.Format("{0}-{1}-{2} {3}", arrDateTimeRange[0], arrDateTimeRange[1], arrDateTimeRange[2], arrDateTimeRange[3]);
                if (!DateTime.TryParseExact(strStart, "yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTimeRangeObj.Start))
                {
                    dateTimeRangeObj.Start = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                }
                strEnd = string.Format("{0}-{1}-{2} {3}", arrDateTimeRange[4], arrDateTimeRange[5], arrDateTimeRange[6], arrDateTimeRange[7]);
                if (!DateTime.TryParseExact(strEnd, "yyyy-MM-dd HH:mm", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out dateTimeRangeObj.End))
                {
                    dateTimeRangeObj.End = new DateTime(1970, 1, 1, 0, 0, 0, 0);
                }
            }
            return dateTimeRangeObj;
        }
        #region GetPicThumbnail  
        /// <summary>
        /// GetPicThumbnail FOR NET CORE
        /// </summary>
        /// <param name="sFile"></param>
        /// <param name="dFile"></param>
        /// <param name="dHeight"></param>
        /// <param name="dWidth"></param>
        /// <param name="flag">40 to 100</param>
        /// <returns></returns>
        public static bool GetPicThumbnail(string sFile, string dFile, int dHeight, int dWidth, int flag)
        {
            System.Drawing.Image iSource = System.Drawing.Image.FromFile(sFile);
            ImageFormat tFormat = iSource.RawFormat;
            int sW = 0, sH = 0;
            Bitmap ob;
            //按比例缩放 
            Size tem_size = new Size(iSource.Width, iSource.Height);
            if (tem_size.Width > dHeight || tem_size.Height > dWidth) //将**改成c#中的或者操作符号
            {
                if ((tem_size.Width * dHeight) > (tem_size.Height * dWidth))
                {
                    sW = dWidth;
                    sH = (dWidth * tem_size.Height) / tem_size.Width;
                    ob = new Bitmap(sW, sH);
                }
                else
                {
                    sH = dHeight;
                    sW = (dHeight * tem_size.Width) / tem_size.Height;
                    ob = new Bitmap(sW, sH);
                }
            }
            else
            {
                sW = tem_size.Width;
                sH = tem_size.Height;
                ob = new Bitmap(sW, sH);
            }

            Graphics g = Graphics.FromImage(ob);
            g.Clear(Color.White); //g.Clear(Color.Transparent); 透明
            g.CompositingQuality = CompositingQuality.HighQuality;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.DrawImage(iSource, new Rectangle(0, 0, sW, sH), 0, 0, iSource.Width, iSource.Height, GraphicsUnit.Pixel);
            g.Dispose();
            //以下代码为保存图片时，设置压缩质量 
            EncoderParameters ep = new EncoderParameters();
            long[] qy = new long[1];
            qy[0] = flag;//设置压缩的比例1-100 
            EncoderParameter eParam = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, qy);
            ep.Param[0] = eParam;
            try
            {
                ImageCodecInfo[] arrayICI = ImageCodecInfo.GetImageEncoders();
                ImageCodecInfo jpegICIinfo = null;
                for (int x = 0; x < arrayICI.Length; x++)
                {
                    if (arrayICI[x].FormatDescription.Equals("JPEG"))
                    {
                        jpegICIinfo = arrayICI[x];
                        break;
                    }
                }
                if (jpegICIinfo != null)
                {
                    ob.Save(dFile, jpegICIinfo, ep);//dFile是压缩后的新路径 
                }
                else
                {
                    ob.Save(dFile, tFormat);
                }
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                iSource.Dispose();
                ob.Dispose();

            }
        }
        #endregion
        public static string GetProperties<T>(T t)
        {
            string tStr = string.Empty;
            if (t == null)
            {
                return tStr;
            }
            System.Reflection.PropertyInfo[] properties = t.GetType().GetProperties(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public);

            if (properties.Length <= 0)
            {
                return tStr;
            }
            foreach (System.Reflection.PropertyInfo item in properties)
            {
                string name = item.Name;
                object value = item.GetValue(t, null);
                if (item.PropertyType.IsValueType || item.PropertyType.Name.StartsWith("String"))
                {
                    tStr += string.Format("{0}:{1},", name, value);
                }
                else
                {
                    GetProperties(value);
                }
            }
            return tStr;
        }

        /// <summary>
        /// Judge Has Property 
        /// </summary>
        /// <param name="PropertyName">PropertyName</param>
        /// <param name="o">Object</param>
        /// <returns></returns>
        public static bool JudgeHasProperty(string PropertyName, Object o)
        {
            if (o == null)
            {
                o = new { };
            }
            PropertyInfo[] p1 = o.GetType().GetProperties();
            bool b = false;
            foreach (PropertyInfo pi in p1)
            {
                if (pi.Name.ToLower() == PropertyName.ToLower())
                {
                    b = true;
                }
            }
            return b;
        }

        /// <summary>
        /// 模分钟数区时间
        /// </summary>
        /// <returns></returns>
        public static DateTime getQuarterDateTime()
        {
            DateTime dt = DateTime.Now;
            int mins = DateTime.Now.Minute;
            int mod = mins % 15;
            mod = -mod;
            dt = DateTime.Now.AddMinutes(mod);
            return dt;
        }
        /// <summary>
        /// 模分钟数区时间
        /// </summary>
        /// <param name="mins_mod">%分钟数</param>
        /// <returns></returns>
        public static DateTime GetQuarterDateTime(int mins_mod)
        {
            if (mins_mod < 1)
            {
                return DateTime.Now;
            }
            DateTime dt = DateTime.Now;
            int mins = DateTime.Now.Minute;
            int mod = mins % mins_mod;
            mod = -mod;
            dt = DateTime.Now.AddMinutes(mod);
            return dt;
        }
    }
    public class DateTimeRangeObj
    {
        public DateTime Start;
        public DateTime End;
    }
    public enum PictureSize
    {
        IsNotPict = 0, s48X48 = 1, s60X60 = 2, s100X100 = 3, s160X160 = 4, s230X230 = 5, s250X250 = 6, s310X310 = 7, s350X350 = 8, s600X600 = 9
    }
    /// <summary>
    /// Thumbnail file name suffix
    /// </summary>
    public class PictureSuffix
    {
        public static string ReturnSizePicUrl(string PicUrl, PictureSize pictureSize)
        {
            if (PicUrl.ToLower().IndexOf("gif") != -1)
            {
                return PicUrl + pictureSize + ".gif";
            }
            if (PicUrl.ToLower().IndexOf("png") != -1)
            {
                return PicUrl + pictureSize + ".png";
            }
            if (PicUrl.ToLower().IndexOf("jpeg") != -1)
            {
                return PicUrl + pictureSize + ".jpeg";
            }
            return PicUrl + pictureSize + ".jpg";

        }
    }

    public enum AccountMode
    {
        EMAIL = 0, MOBILE = 1, NAME = 2, NOT_MATCH= 3
    }
    
}
