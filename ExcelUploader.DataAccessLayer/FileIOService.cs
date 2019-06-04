using ExcelUploader.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUploader.DataAccessLayer
{
    public class FileToDBService
    {
        // This class Handles any CRUD operation will be ran on the DB
        
        public string UploadPath { get; set; }
        public FileToDBService()
        {

            UploadPath = "~/Uploads/";
        }
        public string GetUploadPath()
        {
            return UploadPath;
        }

        public bool CreateNewTable(string sqlStatement)
        {
            // Here we pass the sql cmd we wish to run on DB 
            using (ExcelUploaderContext ctx = new ExcelUploaderContext())
            {
                try
                {
                    int result = ctx.Database.ExecuteSqlCommand(sqlStatement);
                    return true;
                }
                catch (Exception )
                {

                    

                    return false;
                }
            }
        }
        public bool DropSqlTable(string TableName)
        {
            using (ExcelUploaderContext ctx = new ExcelUploaderContext())
            {
                try
                {
                    int result = ctx.Database.ExecuteSqlCommand("DROP TABLE "+TableName);
                    return true;
                }
                catch (Exception )
                {

                    
                    return false;
                }
            }
        }
        public bool PopulateTableData(string excelConString, ExcelFileSchema Schema, string sqlConString)
        {
            // step 1 -- Here we pass excel connection string and connect to it then fill a newly created data table with it

            DataTable dt = new DataTable();
            OleDbConnection excelConn = new OleDbConnection(excelConString);
            excelConn.Open();
            DataTable FileSchema = excelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null); // no restrictions


            string sheetName = FileSchema.Rows[0]["TABLE_NAME"].ToString();
            string selectCmd = "select * from [" + sheetName + "]";

            OleDbDataAdapter SheetAdapter = new OleDbDataAdapter(selectCmd, excelConn);
            
            SheetAdapter.Fill(dt);//  the data is filled from the excel file to the data table that we will use to fill the DB table with.

            // step 2 --- after thr data is filled into the D.T. then special characters in column names are handled;
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = this.HandleSpecialChars(dt.Columns[i].ColumnName).NewString;
            }
            excelConn.Close();

            // step 3 -- then here we get  connect to the db , and map the excel tabel to the db table

            SqlConnection con = new SqlConnection(sqlConString);

            try
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {
                    sqlBulkCopy.DestinationTableName = "dbo." + Schema.TableName;
                    var columns = Schema.ColumnsNames;
                    for (int i = 0; i < columns.Count; i++)
                    {
                        sqlBulkCopy.ColumnMappings.Add(columns[i], columns[i]); // here we had to loop over the columns gotten from the excel file and then map each column name to a column with the same name in the newly created database table

                    }
                    con.Open();
                    sqlBulkCopy.WriteToServer(dt); // filling the DB table with data from the data table;
                    con.Close();
                    return true;
                }
            }
            catch (Exception)
            {

                

                return false;
            }
        }

        public DataTable GetTableData(string tableName, string sqlConnectionString)
        {
            // here we get the data from DB table and fill a data table with it;

            DataTable dt = new DataTable();
            string query = "select * from " + tableName;
            try
            {
                SqlConnection con = new SqlConnection(sqlConnectionString);
                SqlCommand cmd = new SqlCommand(query, con);
                con.Open();

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                con.Close();
                da.Dispose();
            }
            catch (Exception ex)
            {
                    

            }

            return dt;
        }
        public FileNameValidation HandleSpecialChars(string text)
        {
            // this method is also repeated in the Business logic layer Project service class to not
            // make circular reference between it and this project
            FileNameValidation fileNameValidation = new FileNameValidation();


            string specialChars = "[*'\",_&#^@ ]-=+~()!<>$;";

            for (int i = 0; i < specialChars.Length; i++)
            {
                bool hasSpecChar = text.Contains(specialChars[i]);

                if (hasSpecChar)
                {
                    // If a special character is found we replace it with an _ .

                    fileNameValidation.HasSpecChar = true;
                    text = text.Replace(specialChars[i], '_');
                }
            }
            fileNameValidation.NewString = text;

            return fileNameValidation;
        }
    }

}
