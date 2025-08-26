using System; 
using System.Collections.Generic; 
using System.Drawing; 
using System.IO;
using System.Linq;
using System.Net;
using System.Text;   
using System.Configuration;

using System.Text.RegularExpressions;
using Newtonsoft.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Data;
using ExcelDataReader;

namespace MailJobService
{
    public partial class MailCommonBase
    {
        #region   EXCEL READER

        public static DataTable ReadExcel(string filePath)
        {  
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            { 
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                { 
                    if(reader==null)
                    {
                        return null;
                    }
                    do { while (reader.Read()) { } } while (reader.NextResult());

                    var configuration = new ExcelDataSetConfiguration { ConfigureDataTable = tableReader => new ExcelDataTableConfiguration { UseHeaderRow = true } };
                    try
                    {
                        var result = reader.AsDataSet(configuration);

                        return result.Tables[0];
                    }
                    catch
                    {
                        return null;
                    }
                   
                }
                
            }
        }
        #endregion  EXCEL READER 
        //--------------------------------------------------------------------------------------
         
    }
}
