using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Validation;
using System.Data.OleDb;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using ImportToExcel.Models;
using ImportToExcel.Data;
using LinqToExcel;
using System.Data.SqlClient;

namespace ImportToExcel.Controllers
{
    public class ImportToExcelController : Controller
    {
        // GET: ImportToExcel

        private DemoEntities db = new DemoEntities();
        public ActionResult Index()
        {
            return View();
        }
        public FileResult DownloadExcel()
        {
            string path = "/Document/AspirantsData.xlsx";
            return File(path, "application/vnd.ms-excel", "AspirantsData.xlsx");
        }

        [HttpPost]
        public JsonResult UploadExcel(ImportToExcelController importToExcel, HttpPostedFileBase
       FileUpload)
        {
            List<string> data = new List<string>();
            if (FileUpload != null)
            {
                // tdata.ExecuteCommand("truncate table OtherCompanyAssets");
                if (FileUpload.ContentType == "application/vnd.ms-excel" || FileUpload.ContentType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                {
                    string filename = FileUpload.FileName;
                    string targetpath = Server.MapPath("~/Document/");
                    FileUpload.SaveAs(targetpath + filename);
                    string pathToExcelFile = targetpath + filename;
                    var connectionString = "";
                    if (filename.EndsWith(".xls"))
                    {
                        connectionString =
                       string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties = Excel 8.0; ", pathToExcelFile);
                    }
                    else if (filename.EndsWith(".xlsx"))
                    {
                        connectionString =
                       string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties =\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", pathToExcelFile);
                    }
                    var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]",
                   connectionString);
                    var ds = new DataSet();
                    adapter.Fill(ds, "ExcelTable");
                    DataTable dtable = ds.Tables["ExcelTable"];
                    string sheetName = "Sheet1";
                    var excelFile = new ExcelQueryFactory(pathToExcelFile);
                    var artistAlbums = from a in
                   excelFile.Worksheet<tblAspirantsData>(sheetName)
                                       select a;
                    foreach (var a in artistAlbums)
                    {
                        try
                        {
                            if(a.AspirantName != "" && a.Degree != "")
                            {
                                tblAspirantsData TU = new tblAspirantsData();
                                //TU.Ids = a.Id.ToString("i");
                                //TU.AspirantIds = a.AspirantId.ToString("l");
                                TU.AspirantName = a.AspirantName;
                                TU.Degree = a.Degree;
                               // TU.Markss = a.Marks.ToString("mm");
                                //TU.Dstring = a.PassoutYear.ToString("yyyy");
                                db.tblAspirantsDatas.Add(TU);
                                db.SaveChanges();
                            }
                            else
                            {
                                data.Add("<ul>");
                                if (a.AspirantIds == "" || a.AspirantIds == null)
                                    data.Add("<li> name is required</li>");
                                if (a.AspirantName == "" || a.AspirantName == null)
                                    data.Add("<li> name is required</li>");
                                if (a.Degree == "" || a.Degree == null)
                                   data.Add("<li> MOBILE is required</li>");
                                if (a.Markss == "" || a.Markss == null)
                                   data.Add("<li>ContactNo is required</li>");
                                if (a.Dstring == "" || a.Dstring == null)
                                    data.Add("<li>ContactNo is required</li>");
                                data.Add("</ul>");
                                data.ToArray();
                                return Json(data, JsonRequestBehavior.AllowGet);
                            }
                        }
                        catch (DbEntityValidationException ex)
                        {
                            foreach (var entityValidationErrors in
                           ex.EntityValidationErrors)
                            {
                                foreach (var validationError in
                               entityValidationErrors.ValidationErrors)
                                {
                                    Response.Write("Property: " +
                                   validationError.PropertyName + " Error: " + validationError.ErrorMessage);
                                }
                            }
                        }
                    }
                    //deleting excel file from folder
                    if ((System.IO.File.Exists(pathToExcelFile)))
                    {
                        System.IO.File.Delete(pathToExcelFile);
                    }
                    return Json("success", JsonRequestBehavior.AllowGet);
                }
                else
                {
                    //alert message for invalid file format
                    data.Add("<ul>");
                    data.Add("<li>Only Excel file format is allowed</li>");
                    data.Add("</ul>");
                    data.ToArray();
                    return Json(data, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                data.Add("<ul>");
                if (FileUpload == null) data.Add("<li>Please choose Excelfile </ li > ");
                data.Add("</ul>");
                data.ToArray();
                return Json(data, JsonRequestBehavior.AllowGet);
            }
        }

    }
}