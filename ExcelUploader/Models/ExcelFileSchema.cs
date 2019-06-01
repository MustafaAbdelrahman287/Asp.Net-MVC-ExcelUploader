using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelUploader.Models
{
    public class ExcelFileSchema
    {
        public string  TableName { get; set; }
        public string SheetName { get; set; }
        public List<string> ColumnsNames { get; set; }


        public ExcelFileSchema()
        {
            ColumnsNames = new List<string>();
        }
    }
}