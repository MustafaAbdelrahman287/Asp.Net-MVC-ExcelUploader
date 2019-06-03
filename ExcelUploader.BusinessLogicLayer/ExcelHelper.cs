using ExcelUploader.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using ExcelUploader.DataAccessLayer;
using System.Web;
using System.IO;
namespace ExcelUploader.BusinessLogicLayer
{
    public class ExcelService
    {
        private readonly FileToDBService _DbService;
        private readonly ConnectionStringHelper _ConnectionStrHelper;

        private string[] _allowedExtentions { get; set; }

        public ExcelService() // constructor
        {
            _ConnectionStrHelper = new ConnectionStringHelper();

            _allowedExtentions = new[] { ".xsl", ".xlsx" };

            _DbService = new FileToDBService();
        }
        public FileValidation ValidateFile(HttpPostedFileBase postedFile, string path)
        {
            string fileName;

            FileValidation fileValidation = new FileValidation();

            fileValidation.FileExists = this.CheckIfFileExists(path + postedFile.FileName);

            if (!fileValidation.FileExists)
            {

                var fileExtension = Path.GetExtension(postedFile.FileName).ToLower();
                fileName = Path.GetFileNameWithoutExtension(path);

                if (_allowedExtentions.Contains(fileExtension))
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    fileValidation.hasError = false;
                    fileValidation.Message = "OK! Your File Is Uploaded Successfully!";
                }
                else
                {
                    fileValidation.hasError = true;
                    fileValidation.Message = "Please Select Excel Files Only!";
                }
            }
            else
            {
                fileValidation.hasError = true;
                fileValidation.Message = "A file with the same name already exists!";
            }
            return fileValidation;
        }

        public bool SaveFileToDesk(HttpPostedFileBase postedFile, string path)
        {
            try
            {
                postedFile.SaveAs(path);
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }
        public ExcelFileSchema GetExcelFileSchema(string connectionString, string fileName)
        //This function can be enhanced to access the schema of multiple sheets within
        // a single excel file.
        {
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();
            DataTable FileSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null); // no restrictions


            string sheetName = FileSchema.Rows[0]["TABLE_NAME"].ToString();
            string selectCmd = "select * from [" + sheetName + "]";

            OleDbDataAdapter SheetAdapter = new OleDbDataAdapter(selectCmd, conn);
            SheetAdapter.Fill(FileSchema);

            conn.Close();
            ExcelFileSchema schema = new ExcelFileSchema();
            foreach (DataRow row in FileSchema.Rows)
            {
                string columnName = row["Column_Name"].ToString();
                if (!string.IsNullOrEmpty(columnName))
                {
                    schema.ColumnsNames.Add(columnName);

                }
            }
            schema.SheetName = sheetName;
            schema.TableName = fileName;
            return schema;
        }

        public bool SaveFileToDB(HttpPostedFileBase postedFile, string excelPath, string dbPath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelPath);

            string connString = _ConnectionStrHelper.GetExcelConnectionString(Path.GetExtension(postedFile.FileName), excelPath);

            var fileSchema = this.GetExcelFileSchema(connString, fileNameWithoutExtension);

            string sqlCreateStatment = this.GenerateSQLCreateStatement(fileSchema);

            string sqlConString = _ConnectionStrHelper.GetSQLConnectionString(dbPath);
            bool tableCreated = _DbService.CreateNewTable(sqlCreateStatment);

            bool dataAdded = _DbService.PopulateTableData(connString, fileSchema, sqlConString);
            if (tableCreated && dataAdded)
            {
                return true;
            }
            else
            {
                return false;

            }
        }

        public bool UpdateFileInDB(HttpPostedFileBase postedFile, string excelPath, string dbPath)
        {
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(excelPath);
            bool oldTableDropped = _DbService.DropSqlTable(fileNameWithoutExtension);
            if (oldTableDropped)
            {
                string connString = _ConnectionStrHelper.GetExcelConnectionString(Path.GetExtension(postedFile.FileName), excelPath);

                var fileSchema = this.GetExcelFileSchema(connString, fileNameWithoutExtension);

                string sqlCreateStatment = this.GenerateSQLCreateStatement(fileSchema);

                string sqlConString = _ConnectionStrHelper.GetSQLConnectionString(dbPath);
                bool tableCreated = _DbService.CreateNewTable(sqlCreateStatment);

                bool dataAdded = _DbService.PopulateTableData(connString, fileSchema, sqlConString);
                if (tableCreated && dataAdded)
                {
                    return true;
                }
                else
                {
                    return false;

                }
            }
            else return false;
        }
        public string GetUploadPath()
        {
            return _DbService.GetUploadPath();
        }

        public bool DeleteFile(string path)
        {
            try
            {
                File.Delete(path);
                return true;
            }
            catch (Exception ex)
            {

                return false;
            }
        }
        private string GenerateSQLCreateStatement(ExcelFileSchema excelFileSchema)
        {
            var columns = excelFileSchema.ColumnsNames;
            string columnsCommandPart = "";

            for (int i = 0; i < columns.Count; i++)
            {
                // if it's not the last column name we still will generate comma at the end
                // else will not generate the comma to avoid errors when running the script on DB.
                if (columns[i] != columns.Last())
                {
                    columnsCommandPart += columns[i] + " nvarchar(max) ,";
                }
                else
                {
                    columnsCommandPart += columns[i] + " nvarchar(max)";
                }
            }
            string command = string.Format(@"CREATE TABLE {0} ( {1} )", excelFileSchema.TableName, columnsCommandPart);

            return command;
        }

        private bool CheckIfFileExists(string path)
        {

            return File.Exists(path);
        }


    }
}
