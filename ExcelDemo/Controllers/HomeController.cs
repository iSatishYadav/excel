using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDemo.Models;
using System;
using System.Collections.Generic;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace ExcelDemo.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        [HttpPost]
        public ActionResult About(FileModel model)
        {
            if (ModelState.IsValid)
            {
                var package = Package.Open(model.File.InputStream, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                using (var spreadSheet = SpreadsheetDocument.Open(package))
                {
                    var workbookPart = spreadSheet.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var stringBuilder = new StringBuilder();

                    #region Not for larger files
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    foreach (var row in sheetData.Elements<Row>())
                    {
                        foreach (var cell in row.Elements<Cell>())
                        {
                            var cellValue = cell.DataType;
                            if (cellValue.Value == CellValues.SharedString)
                            {
                                int id = -1;
                                SharedStringItem text;
                                text =  GetSharedStringValue(workbookPart, cell, ref id);
                                if(text!=null)
                                {
                                    stringBuilder.Append(text.Text.Text);
                                    continue;
                                }
                            }
                            stringBuilder.Append(cell.CellValue.Text);
                        }
                    }
                    #endregion

                    #region For Larger sheets
                    //var openXMLReader = OpenXmlReader.Create(workbookPart);
                    //while (openXMLReader.Read())
                    //{
                    //    if(openXMLReader.ElementType == typeof(CellValue))
                    //    {
                    //        stringBuilder.Append(openXMLReader.GetText());
                    //    }
                    //}
                    #endregion

                    ViewBag.Message = stringBuilder.ToString();
                    return View();
                }
            }
            return View(model);
        }

        public static SharedStringItem GetSharedStringValue( WorkbookPart workbookPart, Cell cell, ref int id)
        {
            if (int.TryParse(cell.InnerText, out id))
            {
                var text = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                return text;
            }
            return null;
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}