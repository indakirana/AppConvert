using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using AppConvert.Models;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace AppConvert.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {
        return View();
    }

    public IActionResult ConvertExcelToPdf()
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //load an existing file
                FileStream excelStream = new FileStream("Data/sample_excel.xls", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();
                
                // The first worksheet object in the worksheets collection is accessed
                IWorksheet worksheet = workbook.Worksheets[0];
 
                // Set display mode
                worksheet.PageSetup.Orientation = ExcelPageOrientation.Landscape;

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Create the MemoryStream to save the converted PDF.      
                MemoryStream pdfStream = new MemoryStream();

                //Save the converted PDF document to MemoryStream.
                pdfDocument.Save(pdfStream);
                pdfStream.Position = 0;

                //Download PDF document in the browser.
                return File(pdfStream, "application/pdf", "Sample.pdf");
            }
        }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
