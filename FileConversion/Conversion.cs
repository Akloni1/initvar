using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
//using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileConversion
{
    public static class Conversion
    {

        public static MemoryStream ReadAllBytesToMemoryStream(string path)
        {
            byte[] buffer = File.ReadAllBytes(path);
            var destStream = new MemoryStream(buffer.Length);
            destStream.Write(buffer, 0, buffer.Length);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static MemoryStream CopyFileStreamToMemoryStream(string path)
        {
            using FileStream sourceStream = File.OpenRead(path);
            var destStream = new MemoryStream((int)sourceStream.Length);
            sourceStream.CopyTo(destStream);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static FileStream CopyFileStreamToFileStream(string sourcePath, string destPath)
        {
            using FileStream sourceStream = File.OpenRead(sourcePath);
            FileStream destStream = File.Create(destPath);
            sourceStream.CopyTo(destStream);
            destStream.Seek(0, SeekOrigin.Begin);
            return destStream;
        }

        public static FileStream CopyFileAndOpenFileStream(string sourcePath, string destPath)
        {
            File.Copy(sourcePath, destPath, true);
            return new FileStream(destPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }

        public static void DoWorkCloningOpenXmlPackage()
        {
          //  using WordprocessingDocument sourceWordDocument1 = WordprocessingDocument.Open(@"C:\Users\Михаил\Desktop\macros.docx", false);
            using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(@"C:\Users\Михаил\Desktop\test.docx", false);
             using var wordDocument = (WordprocessingDocument)sourceWordDocument.Clone(@"C:\Users\Михаил\Desktop\rr.docx", true);
          //  using WordprocessingDocument sourceWordDocument = WordprocessingDocument.Open(@"/home/miharp/Загрузки/test.docx", false);
           // using var wordDocument = (WordprocessingDocument)sourceWordDocument.Clone(@"/home/miharp/Загрузки/testQ.docx", true);

            ChangeWordprocessingDocument(wordDocument);
        }

        private static void ChangeWordprocessingDocument(WordprocessingDocument wordDocument)
        {
            int count = 0;
            Body body = wordDocument.MainDocumentPart.Document.Body;
            //  Text text = body.Descendants<Text>().First();
            List<Text> text = body.Descendants<Text>().ToList();

            foreach (Text item in text)
            {
                if (item.Text.Contains("название организации"))
                {
                    text[count].Text = text[count].Text.Replace("название организации", "Ланит");
                }
                count++;
            }
          //  text[0].Text = "Не согласовано";
        }
        public static void ConversionFile()
        {
            /*Body gg;
            string fileName = @"C:\Users\Михаил\Desktop\test.docx";
            

            string fileSrc = @"C:\Users\Михаил\Desktop\test.docx";
            string fileRes = @"C:\Users\Михаил\Desktop\rrrrrrrrrrrrrrrrrr.docx";
            using (var docSrc = WordprocessingDocument.Open(fileSrc, false))
            using (var docRes = WordprocessingDocument
              .Create(fileRes, WordprocessingDocumentType.Document))
            {
                docRes.AddMainDocumentPart();
                var sb = new StringBuilder();
                
                using (var stream = new StreamReader(docSrc.MainDocumentPart.GetStream()))
                sb.Append(stream.ReadToEnd());
              //  var date = DateTime.Now;
               // sb.Replace("Number", "123");
               // sb.Replace("DateS", date.ToShortDateString());
                using (var stream = new StreamWriter(docRes.MainDocumentPart.GetStream(FileMode.Create)))
                    stream.Write(sb);
            }*/





           

        }
    }
}
