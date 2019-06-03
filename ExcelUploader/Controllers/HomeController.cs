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
using ExcelUploader.Models.ViewModels;

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
            var model = new HomeViewModel();

            return View(model);
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase postedFile)
        {
            var homeVM = new HomeViewModel();
            // handle special chars here
            if (postedFile != null)
            {

                string path = Server.MapPath(_excelService.GetUploadPath()); // Upload Path
                string dbPath = Server.MapPath(_connectionStringHelper.GetDBPath()); // db Path
                string fileName = postedFile.FileName;
                // handling speacial characters
                var nameValidation = _excelService.HandleSpecialChars(postedFile.FileName);
                if (nameValidation.HasSpecChar)
                {
                    fileName = nameValidation.NewString; // if there are spec chars it's now replaced in file name 
                }
                homeVM.FileValidation = _excelService.ValidateFile(postedFile, path, fileName);

                if (!homeVM.FileValidation.hasError)
                {
                    
                    bool savedToDesk = _excelService.SaveFileToDesk(postedFile, path + fileName);
                    
                    if (savedToDesk)
                    {
                        
                        bool savedToDB = _excelService.SaveFileToDB(postedFile, path + fileName, dbPath);
                        // passed file , file path on desk and DbFilePath to SaveToDB Method
                        if (!savedToDB)
                        {
                            homeVM.FileValidation.Message = "Oops! Something went wrong.. the file maybe in use!";
                            homeVM.FileValidation.hasError = true;
                        }
                    }
                    else
                    {
                        homeVM.FileValidation.Message = "Oops! Something went wrong.. the file maybe in use!";
                        homeVM.FileValidation.hasError = true;
                    }
                }
                if (homeVM.FileValidation.FileExists)
                {
                   homeVM = this.Update(postedFile, fileName, path, dbPath);
                    
                }

                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(path + fileName);
                homeVM.FileName = fileNameWithoutExt;
                homeVM.FileData = _excelService.GetTableData(fileNameWithoutExt, dbPath);
            }
            return View(homeVM);
        }
        
        public HomeViewModel Update(HttpPostedFileBase postedFile, string fileName , string uploadPath, string dbPath)
        {

            var homeVM = new HomeViewModel();

            bool oldFileDeleted = _excelService.DeleteFile(uploadPath + fileName);

            if (oldFileDeleted)
            {
               // save new file to desk
                    bool savedToDesk = _excelService.SaveFileToDesk(postedFile, uploadPath +fileName);
                    if (savedToDesk)
                    {
                    
                    // note that handling special chars is already handled;

                    bool savedToDB = _excelService.UpdateFileInDB(postedFile, uploadPath + fileName, dbPath);
                        // passed file , file path on desk and DbFilePath to SaveToDB Method
                        if (!savedToDB)
                        {
                            homeVM.FileValidation.Message = "Oops! Something went wrong..";
                            homeVM.FileValidation.hasError = true;
                        }
                    }
                    else
                    {
                        homeVM.FileValidation.Message = "Oops! Something went wrong..";
                        homeVM.FileValidation.hasError = true;
                    }
                }

            homeVM.FileValidation.hasError = false;
            homeVM.FileValidation.FileExists = true;
            homeVM.FileValidation.Message = "File Overwritten Successfully";
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(uploadPath + fileName);
            homeVM.FileName = fileNameWithoutExt;
            homeVM.FileData = _excelService.GetTableData(fileNameWithoutExt, dbPath);
             return homeVM;
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