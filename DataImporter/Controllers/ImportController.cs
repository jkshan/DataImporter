using System;
using System.Configuration;
using System.IO;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using DataImporter.Business;
using DataImporter.Business.Parser;

namespace DataImporter.Controllers
{
    public class ImportController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Excel()
        {
            return View();
        }

        public ActionResult Csv()
        {
            return View();
        }

        [HttpPost]
        public async Task<ActionResult> Excel(HttpPostedFileBase file)
        {
            if(file == null || file.ContentLength == 0)
            {
                ViewBag.ErrorMessage = "Please Select Valid File";
                return View();
            }

            if(!file.FileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                ViewBag.ErrorMessage = "File Type Not Supported";
                return View();
            }

            ViewBag.FileName = Path.GetFileName(file.FileName);

            var ConnectionString = ConfigurationManager.ConnectionStrings["SqlConnection"].ConnectionString;
            var TableName = ConfigurationManager.AppSettings["TargetTableName"];
            var ExcelSheetName = ConfigurationManager.AppSettings["ExcelSheetName"];
            var parser = new ExcelParser(file.InputStream, ExcelSheetName);
            var importer = new Importer(parser, ConnectionString, TableName, 10000);

            var dataTable = await importer.Process();

            return View("Success", dataTable);
        }

        [HttpPost]
        public async Task<ActionResult> Csv(HttpPostedFileBase file)
        {
            if(file == null || file.ContentLength == 0)
            {
                ViewBag.ErrorMessage = "Please Select Valid File";
                return View();
            }

            if(!file.FileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
            {
                ViewBag.ErrorMessage = "File Type Not Supported";
                return View();
            }

            ViewBag.FileName = Path.GetFileName(file.FileName);

            var TableName = ConfigurationManager.AppSettings["TargetTableName"];
            var ConnectionString = ConfigurationManager.ConnectionStrings["SqlConnection"].ConnectionString;
            var parser = new CsvParser(file.InputStream);

            var importer = new Importer(parser, ConnectionString, TableName, 10000);

            var errorRecords = await importer.Process();

            //var dataTable = await new CSVParser(file.InputStream).Process(ConnectionString, TableName);

            return View("Success", errorRecords);
        }


    }
}