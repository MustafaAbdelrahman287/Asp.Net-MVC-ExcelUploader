using ExcelUploader.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
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
                filePath = path + fileName;

                if (allowedExtensions.Contains(checkextension))
                {
                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    postedFile.SaveAs(path + postedFile.FileName);
                    CreateNewTable(fileName);
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
            return conString;
        }

        private void CreateNewTable(string tableName)
        {
            //using (ExcelUploaderContext ctx = new ExcelUploaderContext())
            //{
            //    int result = ctx.Database.ExecuteSqlCommand(string.Format(@"CREATE TABLE {0}", tableName));
            //    OleDbDataAdapter ad = new OleDbDataAdapter();
            //}
        }

        private string[] GetColumnNames(HttpPostedFileBase postedFile)
        {
            var strArray = new string[0];
            return strArray;
        }

        private string GenerateSQLCreateStatment()
        {
            return "";
        }
    }



}