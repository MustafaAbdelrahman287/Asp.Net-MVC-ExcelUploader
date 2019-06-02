using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUploader.Models.ViewModels
{
    public class ValidationVM
    {
        public FileValidation FileValidation { get; set; }
        public bool FileExists { get; set; }
        public ValidationVM()
        {
            FileValidation = new FileValidation();
        }
    }
}
