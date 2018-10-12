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
                    string path = Server.MapPath("~/Content/" + excelfile.FileName);
                    if (System.IO.File.Exists(path))
                        System.IO.File.Delete(path);
                    excelfile.SaveAs(path);
                    // Read data from an excel file
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