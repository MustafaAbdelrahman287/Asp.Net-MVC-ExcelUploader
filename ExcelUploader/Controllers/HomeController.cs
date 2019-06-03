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
            var validationResult = new FileValidation();
            // handle special chars here
            if (postedFile != null)
            {
                string path = Server.MapPath(_excelService.GetUploadPath()); // Upload Path
                string dbPath = Server.MapPath(_connectionStringHelper.GetDBPath()); // db Path

                Session["FileName"] = Path.GetFileNameWithoutExtension(path + postedFile.FileName);
                Session["PostedFile"] = postedFile;

                validationResult = _excelService.ValidateFile(postedFile, path);

                if (!validationResult.hasError)
                {
                  bool savedToDesk = _excelService.SaveFileToDesk(postedFile, path + postedFile.FileName);

                    if (savedToDesk)
                    {
                        bool savedToDB = _excelService.SaveFileToDB(postedFile, path + postedFile.FileName, dbPath);
                        // passed file , file path on desk and DbFilePath to SaveToDB Method
                        if (!savedToDB)
                        {
                            validationResult.Message = "Oops! Something went wrong..";
                            validationResult.hasError = true;
                        }
                    }
                    else
                    {
                        validationResult.Message = "Oops! Something went wrong..";
                        validationResult.hasError = true;
                    }
                }
                if (validationResult.FileExists)
                {
                   ActionResult actionResult = this.Update(postedFile, path, dbPath);

                    return actionResult;
                }
            }
            return View(validationResult);
        }
        
        public ActionResult Update(HttpPostedFileBase postedFile , string uploadPath, string dbPath)
        {

            var validationResult = new FileValidation();
            // handle special chars here
            bool oldFileDeleted = _excelService.DeleteFile(uploadPath + postedFile.FileName);
            if (oldFileDeleted)
            {
               // save new file to desk
                    bool savedToDesk = _excelService.SaveFileToDesk(postedFile, uploadPath + postedFile.FileName);
                    if (savedToDesk)
                    {

                        bool savedToDB = _excelService.UpdateFileInDB(postedFile, uploadPath + postedFile.FileName, dbPath);
                        // passed file , file path on desk and DbFilePath to SaveToDB Method
                        if (!savedToDB)
                        {
                            validationResult.Message = "Oops! Something went wrong..";
                            validationResult.hasError = true;
                        }
                    }
                    else
                    {
                        validationResult.Message = "Oops! Something went wrong..";
                        validationResult.hasError = true;
                    }
                }
            validationResult.hasError = false;
            validationResult.Message = "File Overwritten Successfully";
             return View("Index",validationResult);
        }
        public ActionResult ViewFileData(string fileName)
        {
            return null;
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