using ExcelUploader.Models;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelUploader.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
           // ViewBag.Result = new { hasError = false, Message = "" };
            var allowedExtensions = new[] { ".xsl", ".xlsx"};
            var fileValidation = new FileValidation();
            if (postedFile != null)
            {
                var checkextension = Path.GetExtension(postedFile.FileName).ToLower();
                
                if (allowedExtensions.Contains(checkextension))
                {
                    string path = Server.MapPath("~/Uploads/");
                    string fileName = Path.GetFileNameWithoutExtension(path + postedFile.FileName);

                    if (!Directory.Exists(path))
                    {
                        Directory.CreateDirectory(path);
                    }

                    postedFile.SaveAs(path + postedFile.FileName);
                    CreateNewTable(fileName);
                    fileValidation.hasError = false;
                    fileValidation.Message = "OK! Yor File Is Uploaded Successfully!";
                }
                else
                {
                    fileValidation.hasError = true;
                    fileValidation.Message = "Please Select Excel Files Only!";
                }
            }


            return View(fileValidation);
        }


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

        private void CreateNewTable(string tableName)
        {
            using (ExcelUploaderContext ctx = new ExcelUploaderContext())
            {
                int result = ctx.Database.ExecuteSqlCommand(string.Format(@"CREATE TABLE {0}", tableName));
                OleDbDataAdapter ad = new OleDbDataAdapter();
            }
        }
    }
}