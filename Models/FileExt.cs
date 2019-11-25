using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
namespace Import_And_Export_Data_In_Excel_In_ASP.Net_MVC.Models
{
    public class FileExt : ValidationAttribute
    {
        public string Allow;
        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {
            if (value != null)
            {
                string extension = ((System.Web.HttpPostedFileBase)value).FileName.Split('.')[1];
                if (Allow.Contains(extension))
                    return ValidationResult.Success;
                else
                    return new ValidationResult(ErrorMessage);
            }
            else
                return ValidationResult.Success;
        }
    }
}