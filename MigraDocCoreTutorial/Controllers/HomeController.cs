using Microsoft.AspNetCore.Mvc;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.DocumentObjectModel;
using MigraDocCore.Rendering;
using MigraDocCoreTutorial.Models;
using System.Diagnostics;

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
            var document = new Document();

            // With MigraDocCore, we don’t add pages, we add sections (at least one):
            var section = document.AddSection();

            // Adding text is simple:
            section.AddParagraph("Fruits");

            // Store the newly created object for modification:
            Paragraph paragraph = section.AddParagraph();
            paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
            paragraph.Format.Font.Name = "Malgun Gothic";
            paragraph.AddFormattedText("과일 재고", TextFormat.Underline);

            // AddFormattedText also returns an object:
            FormattedText ft = paragraph.AddFormattedText("really fressh!!", TextFormat.Bold);
            ft.Font.Size = 6;

            // Adding empty paragraphs is even simpler:
            section.AddParagraph();

            var table = section.AddTable();
            table.AddColumn(Unit.FromPoint(100));
            table.AddColumn(Unit.FromPoint(100));
            table.AddColumn(Unit.FromPoint(100));
            table.AddColumn(Unit.FromPoint(100));

            Row r;

            // Title
            r = table.AddRow();
            r.Cells[0].AddParagraph("Name");
            r.Cells[1].AddParagraph("Units");
            r.Cells[2].AddParagraph("Amount");
            r.Cells[3].AddParagraph("Price");

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
                }
            });

            // With MigraDocCore you can create PDF or RTF. Just select the appropriate renderer:
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);

            // Pass the document to the renderer:
            pdfRenderer.Document = document;

            // Let the renderer do its job:
            pdfRenderer.RenderDocument();

            // Save the PDF to a file:
            var stream = new MemoryStream();
            pdfRenderer.PdfDocument.Save(stream);

            return File(stream.ToArray(), "application/pdf", "Fruits.pdf");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}