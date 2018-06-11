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

        public static void GetTextFromPdf(Stream pdfIn, out StringBuilder builder)
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
            PRStream pst;
            PdfImageObject pio;
            PdfObject po;

            int n = reader.XrefSize;
            try
            {
                for (int i = 0; i < n; i++)
                {
                    po = reader.GetPdfObject(i);
                    if (po == null || !po.IsStream())
                        continue;

                    pst = (PRStream)po;
                    PdfObject type = pst.Get(PdfName.SUBTYPE);

                    if (type != null && type.ToString().Equals(PdfName.IMAGE.ToString()))
                    {
                        pio = new PdfImageObject(pst);

                        var image = pio.GetDrawingImage();

                        // only add images larger than 50x50 for OCR processing
                        if (image.Height >= 50 && image.Width >= 50)
                        {
                            byte[] imgdata = pio.GetImageAsBytes();
                            MemoryStream memStream = new MemoryStream(imgdata);
                            imgList.Add(memStream);
                        }
                    }
                }
            }
            catch (Exception e) { log.Error(e.Message); }

            return imgList;
        }

        public static List<Stream> ExtractImagesFromPDF(Stream sourcePdf, TraceWriter log)
        {
            List<Stream> imgList = new List<Stream>();
            PdfReader reader = new PdfReader(sourcePdf);
            PRStream pst;
            PdfImageObject pio;
            PdfObject po;

            int n = reader.XrefSize;
            try
            {
                for (int i = 0; i < n; i++)
                {
                    po = reader.GetPdfObject(i);
                    if (po == null || !po.IsStream())
                        continue;

                    pst = (PRStream)po;
                    PdfObject type = pst.Get(PdfName.SUBTYPE);

                    if (type != null && type.ToString().Equals(PdfName.IMAGE.ToString()))
                    {
                        pio = new PdfImageObject(pst);

                        var image = pio.GetDrawingImage();

                        // only add images larger than 50x50 for OCR processing
                        if (image.Height >= 50 && image.Width >= 50)
                        {
                            byte[] imgdata = pio.GetImageAsBytes();
                            MemoryStream memStream = new MemoryStream(imgdata);
                            imgList.Add(memStream);
                        }
                    }
                }
            }
            catch (Exception e) { log.Error(e.Message); }

            return imgList;
        }


        private static PdfObject FindImageInPDFDictionary(PdfDictionary pg)
        {
            PdfDictionary res =
                (PdfDictionary)PdfReader.GetPdfObject(pg.Get(PdfName.RESOURCES));


            PdfDictionary xobj =
              (PdfDictionary)PdfReader.GetPdfObject(res.Get(PdfName.XOBJECT));
            if (xobj != null)
            {
                foreach (PdfName name in xobj.Keys)
                {

                    PdfObject obj = xobj.Get(name);
                    if (obj.IsIndirect())
                    {
                        PdfDictionary tg = (PdfDictionary)PdfReader.GetPdfObject(obj);

                        PdfName type =
                          (PdfName)PdfReader.GetPdfObject(tg.Get(PdfName.SUBTYPE));

                        //image at the root of the pdf
                        if (PdfName.IMAGE.Equals(type))
                        {
                            return obj;
                        }// image inside a form
                        else if (PdfName.FORM.Equals(type))
                        {
                            return FindImageInPDFDictionary(tg);
                        } //image inside a group
                        else if (PdfName.GROUP.Equals(type))
                        {
                            return FindImageInPDFDictionary(tg);
                        }

                    }
                }
            }

            return null;

        }
    }
}
