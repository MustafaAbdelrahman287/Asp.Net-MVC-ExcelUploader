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
                catch (Exception ex)
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
    }
}
