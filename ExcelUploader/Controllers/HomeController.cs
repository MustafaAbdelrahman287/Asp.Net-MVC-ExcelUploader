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
using System.Data.SqlClient;
using ExcelUploader.DataAccessLayer;
using ExcelUploader.BusinessLogicLayer;

namespace ExcelUploader.Controllers
{
    public class HomeController : Controller
    {
        private readonly ExcelService _excelService;
        private readonly ConnectionStringHelper _connectionStringHelper;
        string dbPath;

        public HomeController()
        {
            _excelService = new ExcelService();
            _connectionStringHelper = new ConnectionStringHelper();

        }

        public ActionResult Index()
        {
            var model = new FileValidation();

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            // ViewBag.Result = new { hasError = false, Message = "" };
            var validationResult = new FileValidation();

            if (postedFile != null)
            {

                string path = Server.MapPath(_excelService.GetUploadPath()); // Upload Path
                
                validationResult = _excelService.ValidateFile(postedFile, path);

                if (!validationResult.hasError)
                {
                    try
                    {
                        postedFile.SaveAs(path + postedFile.FileName);

                        dbPath = Server.MapPath(_connectionStringHelper.GetDBPath());

                        _excelService.SaveFileToDB(postedFile, path + postedFile.FileName, dbPath);
                        // passed file , file path on desk and DbFilePath to SaveToDB Method

                    }
                    catch (Exception)
                    {

                        
                    }
                    
                }
            }
            return View(validationResult);
        }

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

    }

}