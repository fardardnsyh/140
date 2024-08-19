using Syncfusion.Compression.Zip;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.OfficeChartToImageConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocLinkConverter.Utility;
using System.Threading.Tasks;

namespace DocLinkConverter.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult UploadFiles()
        {
            // Checking no of files injected in Request object  
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                    //string filename = Path.GetFileName(Request.Files[i].FileName);  

                    HttpPostedFileBase file = Request.Files[0];
                    string fname;

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        fname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        fname = file.FileName;
                    }
                    //WordDocument doc = new WordDocument(file) 
                    //// Get the complete folder path and store the file inside it. 
                    fname = Path.Combine(Server.MapPath("~/Uploads/"), $"{Path.GetFileNameWithoutExtension(fname)}☺{Guid.NewGuid().ToString("N")}{Path.GetExtension(fname)}");
                    file.SaveAs(fname);


                    // Returns message that successfully uploaded  
                    return Json(fname);
                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            else
            {
                return Json("No files selected.");
            }
        }
        [HttpPost]
        public async Task<ActionResult> SaveConvertedPdf(string fname)
        {
            try
            {
                var outputfname = Path.Combine(Server.MapPath("~/Downloads/"), Path.GetFileNameWithoutExtension(fname) + ".pdf");
                WordDocument wordDocument = new WordDocument();
                wordDocument.Settings.SkipIncrementalSaveValidation = true;
                wordDocument.Open(fname, FormatType.Automatic);
                var helper = new Helper();
                var hyperlinks = helper.FindAllHyperlinks(wordDocument);
                foreach (var linkfield in hyperlinks)
                {
                   await helper.RemoveHyperlink(linkfield);
                }
                wordDocument.ChartToImageConverter = new ChartToImageConverter();
                //Creates an instance of the DocToPDFConverter
                DocToPDFConverter converter = new DocToPDFConverter();
                //Converts Word document into PDF document
                PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);
                //Saves the PDF file 
                pdfDocument.Save(outputfname);
                //Closes the instance of document objects
                pdfDocument.Close(true);
                wordDocument.Close();
                System.IO.File.Delete(fname);
                return Json(Path.GetFileName(outputfname));
            }
            catch (ZipException ex)
            {
                System.IO.File.Delete(fname);
                return Json("File is corrupted.");
            }
        }

        [HttpGet]
        public virtual ActionResult Download(string fname)
        {
            byte[] fileBytes = System.IO.File.ReadAllBytes(Path.Combine(Server.MapPath("~/Downloads/"), fname));
            var filename = fname.Substring(0, fname.IndexOf('☺')) + ".pdf";
            System.IO.File.Delete(Path.Combine(Server.MapPath("~/Downloads/"), fname));
            return File(fileBytes, System.Net.Mime.MediaTypeNames.Application.Octet, filename);
        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}