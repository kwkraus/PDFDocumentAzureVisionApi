using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.ProjectOxford.Vision.Contract;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessing.AzureFunction.Helpers
{
    public static class DocumentProcessor
    {
        public static string ScanImageToString(Stream fileStream)
        {
            VisionHelper vision = new VisionHelper(ConfigurationManager.AppSettings["VisionServiceClientKey"], ConfigurationManager.AppSettings["VisionServiceClientRootApi"]);
            OcrResults results = vision.RecognizeText(fileStream);

            return vision.GetRetrieveText(results);
        }

        public async static Task<MemoryStream> CreateTextDocumentAsync(string content)
        {
            var stream = new MemoryStream();

            var writer = new StreamWriter(stream);
            await writer.WriteAsync(content);
            await writer.FlushAsync();

            stream.Position = 0;
            return stream;
        }

        public static void GetTextFromPdf(byte[] pdfIn, out StringBuilder builder)
        {
            builder = new StringBuilder();
            var reader = new PdfReader(pdfIn);

            for (int i = 1; i <= reader.NumberOfPages; i++)
                builder.Append(PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy()));
        }

        public static List<Stream> ExtractImagesFromPDF(byte[] sourcePdf, TraceWriter log)
        {
            List<Stream> imgList = new List<Stream>();
            PdfReader reader = new PdfReader(sourcePdf);
            PRStream prStream;
            PdfImageObject pdfImgObject;
            PdfObject pdfObject;

            int n = reader.XrefSize;

            try
            {
                for (int i = 0; i < n; i++)
                {
                    pdfObject = reader.GetPdfObject(i);
                    if (pdfObject == null || !pdfObject.IsStream())
                        continue;

                    prStream = (PRStream)pdfObject;
                    PdfObject type = prStream.Get(PdfName.SUBTYPE);

                    if (type != null && type.ToString().Equals(PdfName.IMAGE.ToString()))
                    {
                        pdfImgObject = new PdfImageObject(prStream);

                        var image = pdfImgObject.GetDrawingImage();

                        // only add images larger than 50x50 for OCR processing
                        if (image.Height >= 50 && image.Width >= 50)
                        {
                            byte[] imgdata = pdfImgObject.GetImageAsBytes();
                            MemoryStream memStream = new MemoryStream(imgdata);
                            imgList.Add(memStream);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                log.Error(e.Message);
            }

            return imgList;
        }
    }
}
