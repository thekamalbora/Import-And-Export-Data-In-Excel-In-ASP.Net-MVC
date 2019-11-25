using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace Import_And_Export_Data_In_Excel_In_ASP.Net_MVC.Models
{
    public class ImportExcel
    {
        [Required(ErrorMessage = "Please select file")]
        [FileExt(Allow = ".xls,.xlsx", ErrorMessage = "Only excel file")]
        public HttpPostedFileBase file { get; set; }
    }
}