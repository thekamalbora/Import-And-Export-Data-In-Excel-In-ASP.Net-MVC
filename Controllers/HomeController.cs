using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data.SqlClient;
using System.Data;
using System.Web.UI;
using System.Web.UI.WebControls;
using Newtonsoft.Json;
using System.Configuration;
using Import_And_Export_Data_In_Excel_In_ASP.Net_MVC.Models;
using System.Text;
using System.IO;
using System.Xml;

using System.Data.OleDb;


namespace Import_And_Export_Data_In_Excel_In_ASP.Net_MVC.Controllers
{
    public class HomeController : Controller
    {
        SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString);
        // GET: Home
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public JsonResult EmployeeInsertUpdateData(int EmployeeID, string EmployeeName, int EmployeeAge, string EmployeeProfile, string EmployeeSalary)
        {

            Dictionary<string, string> Dic = new Dictionary<string, string>();
            Dic.Add("Result", "");
            Dic.Add("Status", "0");
            Dic.Add("Focus", "");
            try
            {

                if (EmployeeName == "")
                {
                    Dic["Result"] = "Please Enter EmployeeName..!";
                    Dic["Focus"] = "txtEmployeeName";
                }
                else
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("Sp_EmployeeData_InsertUpdate", con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@EmployeeID", EmployeeID);
                    cmd.Parameters.AddWithValue("@EmployeeName", EmployeeName);
                    cmd.Parameters.AddWithValue("@EmployeeAge", EmployeeAge);
                    cmd.Parameters.AddWithValue("@EmployeeProfile", EmployeeProfile);
                    cmd.Parameters.AddWithValue("@EmployeeSalary", EmployeeSalary);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    con.Close();
                    if (dt.Rows.Count > 0)
                    {
                        Dic["Result"] = dt.Rows[0]["Result"].ToString();
                        Dic["Status"] = dt.Rows[0]["Status"].ToString();
                        GetEmployeeRecord();
                    }
                }
            }
            catch (Exception ex)
            {
                Dic["Result"] = ex.Message;
            }
            return Json(Dic);
        }


        [HttpPost]
        public JsonResult GetEmployeeRecord()
        {

            string data = "";
            con.Open();
            SqlCommand cmd = new SqlCommand("Sp_EmployeeData_Get", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            data = JsonConvert.SerializeObject(dt);
            return Json(data, JsonRequestBehavior.AllowGet);

          
        }
     

  


        public ActionResult ExportToExcel()
        {
            var gv = new GridView();
            con.Open();
            SqlCommand cmd = new SqlCommand("Sp_EmployeeData_Get", con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            gv.DataSource = dt;
            gv.DataBind();
            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=EmployeeRecord"+" "+System.DateTime.Now.ToString("dd MMM yyyy HH:mm:ss").Replace(":", "-") + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter objStringWriter = new StringWriter();
            HtmlTextWriter objHtmlTextWriter = new HtmlTextWriter(objStringWriter);
            gv.RenderControl(objHtmlTextWriter);
            Response.Output.Write(objStringWriter.ToString());
            Response.Flush();
            Response.End();
            return View("Index");
        }


        [HttpPost]
        public ActionResult ImportExcel(ImportExcel importExcel)
        {

            string path = Server.MapPath("~/Content/Upload/" + importExcel.file.FileName);
            importExcel.file.SaveAs(path);

            string excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + ";Extended Properties=Excel 12.0;Persist Security Info=False";
            OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);

            //Sheet Name
            excelConnection.Open();
            string tableName = excelConnection.GetSchema("Tables").Rows[0]["TABLE_NAME"].ToString();
            excelConnection.Close();
            //End

            OleDbCommand cmd = new OleDbCommand("Select * from [" + tableName + "]", excelConnection);
            OleDbCommand cmd1 = new OleDbCommand("Select Count(*) from [" + tableName + "]", excelConnection);
            excelConnection.Open();
            var theCount = (int)cmd1.ExecuteScalar();
            OleDbDataReader dReader;
            dReader = cmd.ExecuteReader();
            SqlBulkCopy sqlBulk = new SqlBulkCopy(ConfigurationManager.ConnectionStrings["DBCS"].ConnectionString);

            //Give your Destination table name
            sqlBulk.DestinationTableName = "EmployeeData";
            //Mappings
          
            
                sqlBulk.ColumnMappings.Add("EmployeeName", "EmployeeName");
                sqlBulk.ColumnMappings.Add("EmployeeAge", "EmployeeAge");
                sqlBulk.ColumnMappings.Add("EmployeeProfile", "EmployeeProfile");
                sqlBulk.ColumnMappings.Add("EmployeeSalary", "EmployeeSalary");
                sqlBulk.WriteToServer(dReader);
            

            
           
            excelConnection.Close();
            
            ViewBag.Result = "Successfully Imported";
            TempData["Count"] = theCount;
            ViewBag.Count = theCount;
            return View("Index");
            
        }


       


    }
}  