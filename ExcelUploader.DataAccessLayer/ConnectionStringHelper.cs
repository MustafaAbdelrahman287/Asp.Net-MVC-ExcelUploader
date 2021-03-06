﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;


namespace ExcelUploader.DataAccessLayer
{
    public class ConnectionStringHelper
    {
        // This class handles Retrieving all connection strings.

        private string _dbPath;
        public ConnectionStringHelper()
        {
            _dbPath = "~/App_Data/ExcelUploader.Models.ExcelUploaderContext.mdf";
            // this is the location that the Entity framework code first mechanizm will anyway create the DB file in by convention if it's not there and name it to the "Assembly.DBContext class name" . 
            
        }

        public string GetExcelConnectionString(string extension, string filePath)
        {
            // in Web.Config file there are two connection strings each for a different format of excel sheets.
            // here we return the corresponding connection string based on format;
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
