using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


namespace ExcelUploader.DataAccessLayer
{
    public class ConnectionStringHelper
    {
        private string _dbPath;
        public ConnectionStringHelper()
        {
            _dbPath = "~/App_Data/ExcelUploader.Models.ExcelUploaderContext.mdf";
        }

        public string GetExcelConnectionString(string extension, string filePath)
        {
            string conString = string.Empty;
            switch (extension)
            {
                case ".xls":
                    conString = ConfigurationManager.ConnectionStrings["Excel2003ConString"].ConnectionString;
                    break;
                case ".xlsx":
                    conString = ConfigurationManager.ConnectionStrings["Excel2007ConString"].ConnectionString;
                    break;
            }
            conString = string.Format(conString, filePath);

            return conString;
        }

        public string GetSQLConnectionString(string dbPath)
        {
            string sqlConSring = ConfigurationManager.ConnectionStrings["DBConnection"].ConnectionString;
            sqlConSring = string.Format(sqlConSring, dbPath);

            return sqlConSring;
        }

        public string GetDBPath()
        {
            return _dbPath;
        }
    }
}
