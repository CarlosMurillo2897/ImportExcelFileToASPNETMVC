using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using ImportExcelFileToASPNETMVC.Models;

namespace ImportExcelFileToASPNETMVC.Controllers
{
    public class UsuarioController : Controller
    {
        private ImportExcelEntities db = new ImportExcelEntities();
        // GET: Usuario
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelfile)
        {
            if (excelfile == null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an excel file<br>";
                return View("Index");
            }
            else
            {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    int fin;
                    string terminacion;

                    if (excelfile.FileName.EndsWith("xlsx")) {
                        fin = excelfile.FileName.Length - 5;
                        terminacion = ").xlsx";
                    }
                    else {
                        fin = excelfile.FileName.Length - 4;
                        terminacion = ").xls";
                    }

                    string name = excelfile.FileName.Substring(0, fin);

                    string path = Server.MapPath("~/Content/" + name +
                                                 "(" + DateTime.Now.Year.ToString() +  "-"
                                                 + DateTime.Now.Month.ToString()    +  "-"
                                                 + DateTime.Now.Day.ToString()      + ")-("
                                                 + DateTime.Now.Hour.ToString()     +  "-"
                                                 + DateTime.Now.Minute.ToString()   +  "-"
                                                 + DateTime.Now.Second.ToString()   +  terminacion);

                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    List<SOGIP_Usuario> usrs = new List<SOGIP_Usuario>();
                    for(int row = 1; row <= range.Rows.Count; row++)
                    {
                        SOGIP_Usuario u = new SOGIP_Usuario();
                        u.id = int.Parse(((Excel.Range)range.Cells[row, 1]).Text);
                        u.cedula = ((Excel.Range)range.Cells[row, 2]).Text;
                        u.contrasena = ((Excel.Range)range.Cells[row, 3]).Text;
                        u.fecha_expiracion = System.DateTime.Parse(((Excel.Range)range.Cells[row, 4]).Text);
                        usrs.Add(u);
                        db.SOGIP_Usuario.Add(u);
                        db.SaveChanges();
                    }
                    ViewBag.ListUsrs = usrs;

                    //¿Necessary to close the process?
                    workbook.Close();

                    System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    foreach (System.Diagnostics.Process p in process)
                    {
                        if (!string.IsNullOrEmpty(p.ProcessName))
                        {
                            try
                            {
                                p.Kill();
                            }
                            catch { }
                        }
                    }


                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect. <br>";
                    return View("Index");
                }
            }
            
        }
    }
}