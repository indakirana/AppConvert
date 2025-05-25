using System.Linq;
using System.Diagnostics;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using AppConvert.Models;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Graphics;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using System.Data;
using System.IO;


namespace AppConvert.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    Db dbop = new Db();

    public IActionResult Index()
    {
        var document = new Document
        {
            PageInfo = new PageInfo { Margin = new MarginInfo(28, 28, 28, 40) }
        };

        var pdfpage = document.Pages.Add();

        Table table = new Table
        {
            ColumnWidths = "25% 25% 25% 25%",
            DefaultCellPadding = new MarginInfo(10, 5, 10, 5),
            Border = new BorderInfo(BorderSide.All, .5f, Color.Black),
            DefaultCellBorder = new BorderInfo(BorderSide.All, .2f, Color.Black),

        };

        DataTable dt = dbop.Getrecord();
        table.ImportDataTable(dt, true, 0, 0);
        document.Pages[1].Paragraphs.Add(table);

        using (var streamout = new MemoryStream())
        {
            document.Save(streamout);
            return new FileContentResult(streamout.ToArray(), "application/pdf")
            {
                FileDownloadName = "table_customer.pdf"
            };
        };
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
