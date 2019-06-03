using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUploader.Models.ViewModels
{
    public class HomeViewModel
    {
        public string FileName { get; set; }
        public FileValidation FileValidation { get; set; }
        public DataTable FileData { get; set; }
        public HomeViewModel()
        {
            FileValidation = new FileValidation();
        }
    }
}
