using ExcelUploader.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelUploader.Controllers
{
    public class HomeController : Controller
    {
        #region About and Contact Actions

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        #endregion

        public ActionResult Index()
        {
            var model = new FileValidation();
            return View(model);
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            // ViewBag.Result = new { hasError = false, Message = "" };
            var allowedExtensions = new[] { ".xsl", ".xlsx" };
            var fileValidation = new FileValidation();
            string filePath = string.Empty;
            string fileName = string.Empty;

            if (postedFile != null)
            {
                var checkextension = Path.GetExtension(postedFile.FileName).ToLower();
                string path = Server.MapPath("~/Uploads/");
                fileName = Path.GetFileNameWithoutExtension(path + postedFile.FileName);
                filePath = path + postedFile.FileName;

                if (allowedExtensions.Contains(checkextension))
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    postedFile.SaveAs(path + postedFile.FileName);
                    fileValidation.hasError = false;
                    fileValidation.Message = "OK! Your File Is Uploaded Successfully!";
                }
                else
                {
                    fileValidation.hasError = true;
                    fileValidation.Message = "Please Select Excel Files Only!";
                }
            }

            string connString = GetConnectionStringForExcelFiles(Path.GetExtension(postedFile.FileName) ,filePath);

            var fileSchema = GetExcelFileSchema(connString);

            string sqlCreateStatment = GenerateSQLCreateStatement(fileName,fileSchema);

            CreateNewTable(sqlCreateStatment);


            return View(fileValidation);
        }

        private string GetConnectionStringForExcelFiles(string extension,  string filePath)
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

        private void CreateNewTable(string sqlStatement)
        {
            using (ExcelUploaderContext ctx = new ExcelUploaderContext())
            {
                int result = ctx.Database.ExecuteSqlCommand(sqlStatement);
            }
        }


        private ExcelFileSchema GetExcelFileSchema(string connectionString) //This function can be enhanced to access the schema of multiple sheets within
                                                                            // a single excel file.
        {
            OleDbConnection conn = new OleDbConnection(connectionString);
            conn.Open();
            DataTable FileSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, null); // no restrictions


            string sheetName = FileSchema.Rows[0]["TABLE_NAME"].ToString();
            string selectCmd = "select * from ["+sheetName+"]";

            OleDbDataAdapter SheetAdapter = new OleDbDataAdapter(selectCmd,conn);
            SheetAdapter.Fill(FileSchema);

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
            return schema;
        }

        private string GenerateSQLCreateStatement(string fileName, ExcelFileSchema excelFileSchema)
        {
            var columns = excelFileSchema.ColumnsNames;
            string columnsCommandPart = "";

            for (int i = 0; i < columns.Count; i++)
            {
                if (columns[i] != columns.Last()) // if it's not the last column name we still will generate comma at the end 
                                                  // else will not generate the comma to avoid erros when running the script on DB.
                {
                    columnsCommandPart += columns[i] + " nvarchar(max) ,";
                }
                else
                {
                    columnsCommandPart += columns[i] + " nvarchar(max)";
                }
            }
            string command = string.Format(@"CREATE TABLE {0} ( {1} )",fileName, columnsCommandPart);

            return command;
        }
    }



}