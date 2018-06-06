using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Azure.WebJobs.Host;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace DocProcessing.AzureFunction.Helpers
{
    public static class OfficeHelper
    {
        public static string GetAllTextFromWordDoc(
            Stream docStream,
            TraceWriter log)
        {
            StringBuilder sb = new StringBuilder();

            try
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(docStream, false))
                {
                    OpenXmlElement body = wordDocument.MainDocumentPart.Document.Body;

                    if (body == null)
                    {
                        return string.Empty;
                    }

                    sb.Append(GetPlainText(body));
                    sb.Append(ExtractTextFromEmbeddedDocuments(wordDocument, log));

                    return sb.ToString();
                }
            }
            catch (FileFormatException fex)
            {
                var errorMsg = $"Word Document is not in OpenXML format or is corrupted";
                log.Error(errorMsg, fex);
                throw new ApplicationException(errorMsg, fex);
            }
            catch(InvalidOperationException ioex)
            {
                var errorMsg = $"Word document is corrupt.  Open in client to fix document.";
                log.Error(errorMsg, ioex);
                throw new ApplicationException(errorMsg, ioex);
            }
            catch (Exception ex)
            {
                var errorMsg = $"Unknown error occurred. Failed to extract all text from Word document.";
                log.Error(errorMsg, ex);
                throw new ApplicationException(errorMsg, ex);
            }
        }

        public static string GetAllTextFromExcelDoc(
            Stream sheetStream,
            TraceWriter log)
        {
            var builder = new StringBuilder();

            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(sheetStream, false))
                {
                    var workbookPart = doc.WorkbookPart;

                    var sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    var sst = sstpart.SharedStringTable;

                    foreach (var worksheetPart in workbookPart.WorksheetParts)
                    {
                        var sheet = worksheetPart.Worksheet;

                        foreach (var row in sheet.Descendants<Row>())
                        {
                            foreach (var c in row.Elements<Cell>())
                            {
                                if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                                {
                                    var ssid = int.Parse(c.CellValue.Text);
                                    var str = sst.ChildElements[ssid].InnerText;
                                    builder.Append(str + " ");
                                }
                                else if (c.CellValue != null)
                                {
                                    builder.Append(c.CellValue.Text + " ");
                                }
                            }
                        }
                    }
                }

                return builder.ToString();
            }
            catch (FileFormatException fex)
            {
                var errorMsg = $"Excel Document is not in OpenXML format or is corrupted";
                log.Error(errorMsg, fex);

                throw new ApplicationException(errorMsg, fex);
            }
            catch (Exception ex)
            {
                var errorMsg = $"Failed to extract all text from Excel document";

                log.Error(errorMsg, ex);
                throw new ApplicationException(errorMsg, ex);
            }
        }

        private static string ExtractTextFromEmbeddedDocuments(
                    WordprocessingDocument element,
                    TraceWriter log)
        {
            var builder = new StringBuilder();

            try
            {
                foreach (var part in element.Parts)
                {
                    var testForEmbedding = part.OpenXmlPart.GetPartsOfType<EmbeddedPackagePart>();

                    foreach (EmbeddedPackagePart embedding in testForEmbedding)
                    {
                        string fileName = embedding.Uri.OriginalString.Split('/').Last();

                        using (var stream = embedding.GetStream())
                        {
                            switch (Path.GetExtension(fileName))
                            {
                                case ".xlsx":
                                    builder.Append(GetAllTextFromExcelDoc(stream, log));
                                    break;

                                case ".docx":
                                    builder.Append(GetAllTextFromWordDoc(stream, log));
                                    break;

                                default:
                                    break;
                            }
                        }
                    }
                }
            }
            catch (ApplicationException aex)
            {
                log.Warning($"Word doc contained embedded documents, but text could not be extracted. Error:{aex.Message}");
            }

            return builder.ToString();
        }


        private static string GetPlainText(OpenXmlElement element)
        {
            StringBuilder PlainTextInWord = new StringBuilder();
            foreach (OpenXmlElement section in element.Elements())
            {
                switch (section.LocalName)
                {
                    // Text
                    case "t":
                        PlainTextInWord.Append(section.InnerText);
                        break;

                    // Carriage return
                    case "cr":
                        break;

                    // Page break
                    case "br":
                        break;

                    // Tab
                    case "tab":
                        break;

                    // Paragraph
                    case "p":
                        PlainTextInWord.Append(GetPlainText(section));
                        break;

                    default:
                        PlainTextInWord.Append(GetPlainText(section));
                        break;
                }
            }

            return PlainTextInWord.ToString();
        }
    }
}