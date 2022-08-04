using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using ImportToExcel.Data;

namespace ImportToExcel.Models
{
    public class ImportToExcelModel
    {
        public int Id { get; set; }
        public int AspirantId { get; set; }
        public string AspirantName { get; set; }
        public string Degree { get; set; }
        public decimal Marks { get; set; }
        public System.DateTime PassoutYear { get; set; }

        public HttpPostedFileBase FileUpload { get; set; }

    }
}