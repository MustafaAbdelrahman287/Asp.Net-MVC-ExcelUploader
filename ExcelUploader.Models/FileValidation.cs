using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelUploader.Models
{
    public class FileValidation
    {
        public bool FileExists { get; set; }
        public bool hasError { get; set; }
        public string  Message { get; set; }
    }
}