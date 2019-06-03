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
        public bool PopulateTableData(string connString, ExcelFileSchema Schema, string sqlConString)
        {

            DataTable dt = new DataTable();
            OleDbConnection excelConn = new OleDbConnection(connString);
            excelConn.Open();
            DataTable FileSchema = excelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null); // no restrictions


            string sheetName = FileSchema.Rows[0]["TABLE_NAME"].ToString();
            string selectCmd = "select * from [" + sheetName + "]";

            OleDbDataAdapter SheetAdapter = new OleDbDataAdapter(selectCmd, excelConn);
            
            SheetAdapter.Fill(dt);

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                dt.Columns[i].ColumnName = this.HandleSpecialChars(dt.Columns[i].ColumnName).NewString;
            }
            excelConn.Close();

            SqlConnection con = new SqlConnection(sqlConString);

            try
            {
                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                {
                    sqlBulkCopy.DestinationTableName = "dbo." + Schema.TableName;
                    var columns = Schema.ColumnsNames;
                    for (int i = 0; i < columns.Count; i++)
                    {
                        sqlBulkCopy.ColumnMappings.Add(columns[i], columns[i]);

                    }
                    con.Open();
                    sqlBulkCopy.WriteToServer(dt);
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
            FileNameValidation fileNameValidation = new FileNameValidation();


            string specialChars = "[*'\",_&#^@ ]-=+~()!<>$;";

            for (int i = 0; i < specialChars.Length; i++)
            {
                bool hasSpecChar = text.Contains(specialChars[i]);
                if (hasSpecChar)
                {
                    fileNameValidation.HasSpecChar = true;
                    text = text.Replace(specialChars[i], '_');
                }
            }
            fileNameValidation.NewString = text;

            return fileNameValidation;
        }
    }

}
