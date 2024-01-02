using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Security.AccessControl;
using System.Text;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System;
using System.Drawing;
using System.IO;
using Tesseract;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Reflection.Emit;
using AngularApiNetCore.Models;
using Freeware;

namespace AngularApiNetCore.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AngularApiController : ControllerBase
    {
        private IConfiguration _configuration;

        public AngularApiController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpPost]
        [Route("GetNotes")]
        public JsonResult GetNotes(UploadFile uploadFile)
        {
            string textOCR = "";
            string ext = Path.GetExtension(@"C:\Users\panup\Downloads\" + uploadFile.PathFile);
            string CountN = "";
            int Num = 0;
            // display image in picture box  
            if (ext != ".pdf" && ext != ".docx" && ext != ".xlsx" && ext != ".csv")
            {
                var ocrengine = new TesseractEngine(@"C:\Users\panup\Downloads\AngularApiNetCore\tessdata", "tha+eng", EngineMode.Default);
                var img = Pix.LoadFromFile(@"C:\Users\panup\Downloads\" + uploadFile.PathFile);
                var res = ocrengine.Process(img);
                textOCR = res.GetText();
                CountN = res.GetText();
                Num = CountN.Length;
            }
            else if (ext == ".docx")
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(@"C:\Users\panup\Downloads\" + uploadFile.PathFile, false))
                {
                    var body = doc.MainDocumentPart.Document.Body;

                    if (body != null)
                    {
                        foreach (var paragraph in body.Elements<Paragraph>())
                        {
                            foreach (var run in paragraph.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                            {
                                foreach (var text in run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>())
                                {
                                    textOCR += text.Text;
                                }
                            }

                            textOCR += Environment.NewLine;
                        }
                    }
                }
            }
            else if (ext == ".csv")
            {
                using (StreamReader reader = new StreamReader(@"C:\Users\panup\Downloads\" + uploadFile.PathFile))
                {
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] fields = line.Split(',');

                        foreach (string field in fields)
                        {
                            textOCR += field + "\t";
                        }

                        textOCR += Environment.NewLine;
                    }
                }
            }
            else if (ext == ".xlsx")
            {
                string textContent = "";

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(@"C:\Users\panup\Downloads\" + uploadFile.PathFile, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    SharedStringTablePart sharedStringTablePart = workbookPart.SharedStringTablePart;

                    if (sharedStringTablePart != null)
                    {
                        // Create a StringBuilder to store the text content
                        StringBuilder stringBuilder = new StringBuilder();

                        // Iterate through all shared strings to extract text
                        foreach (SharedStringItem sharedStringItem in sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
                        {
                            textContent += sharedStringItem.InnerText + "\t";

                            textContent += Environment.NewLine;
                        }
                        // Append shared string text to the StringBuilder
                        //stringBuilder.AppendLine(sharedStringItem.InnerText                       
                        textOCR = textContent;
                    }
                }
                textOCR = textContent;
            }
            else
            {
                var pageText = new StringBuilder();

                //Input File Path
                string input_path = @"C:\Users\panup\Downloads\" + uploadFile.PathFile;
                //Output File Path
                string output_path = @"C:\Users\panup\Downloads\FileUploadTest\PDFOCR" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
                //Tessdata Folder
                string training_data = @"C:\Users\panup\Downloads\AngularApiNetCore\tessdata";

                PdfReader pdf = new PdfReader(input_path);
                PdfDocument pdfDoc = new PdfDocument(pdf);
                int n = pdfDoc.GetNumberOfPages();
                pdf.Close();
                using (IResultRenderer renderer = Tesseract.PdfResultRenderer.CreatePdfRenderer(output_path, training_data, false))
                {
                    using (renderer.BeginDocument("Serachablepdftest"))
                    {
                        for (int i = 1; i <= n; i++)
                        {

                            //GhostscriptWrapper.GeneratePageThumbs(input_path, output_path + i + ".pdf", i, n, 200, 200);
                            byte[] pdfJpg = Pdf2Png.Convert(System.IO.File.ReadAllBytes(input_path), i, 300);
                            System.IO.File.WriteAllBytes(output_path + i + ".jpg", pdfJpg);
                            string configurationFilePath = training_data;
                            string configfile = Path.Combine(training_data, "pdf.ttf");
                            using (TesseractEngine engine = new TesseractEngine(configurationFilePath, "tha+eng", EngineMode.Default, configfile))
                            {
                                using (var img = Pix.LoadFromFile(output_path + i + ".jpg"))
                                {
                                    using (var page = engine.Process(img, "Serachablepdftest"))
                                    {
                                        renderer.AddPage(page);

                                    }
                                }
                                System.IO.File.Delete(output_path + i + ".jpg");
                            }
                        }
                    }
                }

                using (PdfDocument pdfDocument = new PdfDocument(new PdfReader(output_path + ".pdf")))
                {
                    var pageNumbers = pdfDocument.GetNumberOfPages();
                    for (int i = 1; i <= pageNumbers; i++)
                    {
                        LocationTextExtractionStrategy locationTextExtractionStrategy = new LocationTextExtractionStrategy();
                        PdfCanvasProcessor pdfCanvasProcessor = new PdfCanvasProcessor(locationTextExtractionStrategy);
                        pdfCanvasProcessor.ProcessPageContent(pdfDocument.GetPage(i));
                        pageText.AppendLine(locationTextExtractionStrategy.GetResultantText());
                    }
                }
                textOCR = pageText.ToString();
            }

            return new JsonResult(textOCR);
        }
    }
}
