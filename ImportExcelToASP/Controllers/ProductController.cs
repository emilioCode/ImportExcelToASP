using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;
using ImportExcelToASP.Models;

namespace ImportExcelToASP.Controllers
{
    public class ProductController : Controller
    {
        // GET: Product
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Import(HttpPostedFileBase excelFile)
        {
            if (excelFile == null ||excelFile.ContentLength == 0)
            {
                ViewBag.Error = "Por favor seleccione una archivo de excel<br/>";
                return View("Index");
            }else
            {
                if (excelFile.FileName.EndsWith("xls") || excelFile.FileName.EndsWith("xlsx"))
                {
                    //creando el mapeo de camino del archivo
                    string path = Server.MapPath("~/Archivos/" + excelFile.FileName);
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path);
                    }
                    excelFile.SaveAs(path);

                    //Para leer los datos del archivo en Excel
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;

                    List<Product> listProducts = new List<Product>();
                    for (int row = 3; row <= range.Rows.Count; row++)
                    {
                        //var valor = ((Excel.Range)range.Cells[row, 1]).Text;
                        Product product = new Product();
                        product.ProductId = int.Parse(((Excel.Range)range.Cells[row, 1]).Text);
                        product.Nombre = ((Excel.Range)range.Cells[row, 2]).Text;
                        product.Precio = decimal.Parse(((Excel.Range)range.Cells[row, 3]).Text);
                        product.Cantidad = int.Parse(((Excel.Range)range.Cells[row, 4]).Text);
                        listProducts.Add(product);
                    }
                    workbook.Close();
                    application.Quit();
                    


                    ViewBag.ListProducts = listProducts;
                    return View("Success");
                }else
                {
                    ViewBag.Error = "Tipo de archivo incorrecto<br/>";
                    return View("Index");
                }
            }
        }
        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }
    }
}