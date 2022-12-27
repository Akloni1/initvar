using DinkToPdf;
using DinkToPdf.Contracts;
//using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System.IO.Packaging;
using System.Xml.Linq;

namespace FileConversion
{
    public static class Class1
    {




        public static void DocxToPdf()
        {

            byte[] byteArray = File.ReadAllBytes(@"C:\Users\Михаил\Desktop\test10.docx");
            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                {
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = "My Page Title"
                    };
                    XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                    File.WriteAllText(@"C:\Users\Михаил\Desktop\test100.html", html.ToStringNewLineOnAttributes());
                }
            }
        }

        public static void HtmlToPdf()
        {

            var converter = new BasicConverter(new PdfTools());
            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = {
                  ColorMode = ColorMode.Color,
                  Orientation = Orientation.Landscape,
                  PaperSize = PaperKind.A4,
                   Out = @"C:\Users\Михаил\Desktop\test1000.html",
                      },
                Objects = {
                     new ObjectSettings() {
                    PagesCount = true,
                    HtmlContent = File.ReadAllText(@"C:\Users\Михаил\Desktop\test100.html"),
                    WebSettings = { DefaultEncoding = "utf-8" },
                    HeaderSettings = { FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
                    FooterSettings = { FontSize = 9, Right = "Page [page] of [toPage]" }
                           }
                     }

            };

            converter.Convert(doc);

        }





        public static void PdfSharpConvert()
        {
            var htmlContent = String.Format("<body>Hello world: {0}</body>", DateTime.Now);
            var pdfBytes = (new NReco.PdfGenerator.HtmlToPdfConverter()).GeneratePdf(htmlContent);
        }
    }
}