using Microsoft.AspNetCore.Mvc;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.Rendering;
using MigraDocCoreTutorial.Models;
using System.Diagnostics;
using PdfSharpCore.Pdf;
using PdfSharpCore.Drawing;

namespace MigraDocCoreTutorial.Controllers
{
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

        public IActionResult Privacy()
        {
            return View();
        }

        [HttpGet("download")]
        public IActionResult Download()
        {
            #region Dummy Data
            var fruits = new List<Fruit>();
            fruits.Add(new Fruit
            {
                Name = "Apple",
                Unit = "1",
                Amount = 3,
                Price = 9.90m,
            });

            fruits.Add(new Fruit
            {
                Name = "Banana",
                Unit = "a bunch",
                Amount = 5,
                Price = 13.00m
            });

            fruits.Add(new Fruit
            {
                Name = "Blueberry",
                Unit = "500g",
                Amount = 100,
                Price = 10m
            });
            #endregion

            // We start with a new document:
            var fontName = "Malgun Gothic";
            var doc = new PdfDocument();

            var page = doc.AddPage();
            var gfx = XGraphics.FromPdfPage(page);
            gfx.MUH = PdfFontEncoding.Unicode;

            Document document = new Document();
            Unit margin = Unit.FromPoint(50);
            document.DefaultPageSetup.TopMargin = margin;
            document.DefaultPageSetup.BottomMargin = margin;
            document.DefaultPageSetup.LeftMargin = margin;
            document.DefaultPageSetup.RightMargin = margin;

            var docRenderer = new DocumentRenderer(document);

            // With MigraDocCore, we don’t add pages, we add sections (at least one):
            var section = document.AddSection();

            #region 타이틀
            var paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Center;
            paragraph.Format.Font.Name = fontName;
            paragraph.Format.Font.Size = Unit.FromPoint(16);
            //para.Format.Borders.Top.Color = Colors.Black;
            //para.Format.Shading.Color = Colors.LightGray;
            paragraph.Format.Borders.Bottom.Color = Colors.Black;
            paragraph.Format.Borders.DistanceFromTop = Unit.FromPoint(5);
            paragraph.Format.Borders.DistanceFromBottom = Unit.FromPoint(10);

            paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
            paragraph.AddFormattedText($"과일 재고", TextFormat.Bold);
            #endregion


            // AddFormattedText also returns an object:
            FormattedText ft = paragraph.AddFormattedText("really fressh!!", TextFormat.Bold);
            ft.Font.Size = 6;

            paragraph = section.AddParagraph();
            paragraph.Format.Alignment = ParagraphAlignment.Left;
            paragraph.Format.Font.Name = fontName;
            paragraph.Format.Font.Size = Unit.FromPoint(14);
            paragraph.Format.Borders.Top.Color = Colors.LightGray;
            paragraph.Format.Borders.DistanceFromTop = Unit.FromPoint(20);
            paragraph.Format.Borders.DistanceFromBottom = Unit.FromPoint(5);
            paragraph.AddText("오늘의 과일");

            // Adding empty paragraphs is even simpler:
            section.AddParagraph();

            var table = section.AddTable();
            table.Borders.DistanceFromTop = 100;
            table.TopPadding = 5;
            table.BottomPadding = 5;
            table.AddColumn(Unit.FromPoint(125));
            table.AddColumn(Unit.FromPoint(125));
            table.AddColumn(Unit.FromPoint(125));
            table.AddColumn(Unit.FromPoint(125));

            Row r;

            // Title
            r = table.AddRow();
            r.Height = 20;
            r.Cells[0].AddParagraph("Name");
            r.Cells[1].AddParagraph("Units");
            r.Cells[2].AddParagraph("Amount");
            r.Cells[3].AddParagraph("Price");
            foreach (Cell cell in r.Cells)
            {
                cell.Format.Alignment = ParagraphAlignment.Right;
            }
            r.Cells[0].Format.Alignment = ParagraphAlignment.Left;

            // Data
            fruits.ForEach(fruit =>
            {
                r = table.AddRow();
                r.Cells[0].AddParagraph(fruit.Name);
                r.Cells[1].AddParagraph(fruit.Unit);
                r.Cells[2].AddParagraph($"{fruit.Amount}");
                r.Cells[3].AddParagraph($"{fruit.Price}");
                foreach (Cell cell in r.Cells)
                {
                    cell.Format.Font.Name = "Segoe UI";
                    cell.Format.Alignment = ParagraphAlignment.Right;
                }

                r.Cells[0].Format.Alignment = ParagraphAlignment.Left;
            });

            // With MigraDocCore you can create PDF or RTF. Just select the appropriate renderer:
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);

            // Pass the document to the renderer:
            pdfRenderer.Document = document;

            // Let the renderer do its job:
            docRenderer.PrepareDocument();
            docRenderer.RenderPage(gfx, 1);

            // Save the PDF to a file:
            var stream = new MemoryStream();
            doc.Save(stream);

            return File(stream.ToArray(), "application/pdf", "Fruits.pdf");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}