using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.ProjectOxford.Vision;
using Microsoft.WindowsAzure.Storage.Blob;
using Newtonsoft.Json;
using DocProcessing.AzureFunction.Helpers;
using DocProcessing.AzureFunction.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace DocProcessing.AzureFunction
{
    public static class ProcessDocument
    {
        [FunctionName("ProcessDocument")]
        public static async Task Run(
            [QueueTrigger("docprocessingrequests", Connection = "storageresource")]BlobInfo myQueueItem,
            [Blob("originaldocuments/{BlobName}", FileAccess.ReadWrite)] CloudBlockBlob blob,
            TraceWriter log)
        {
            log.Info($"C# Queue trigger function processed: {myQueueItem}");
            log.Info($"Blob info: {blob.Name}");

            ProcessingResult result = null;
            log.Info($"Identifier is ({myQueueItem.Identifier.ToString()}) for this blob");

            try
            {
                string indexingContainer = ConfigurationManager.AppSettings["IndexingContainer"];

                result = await ProcessInputBlobForScanningAsync(
                                    blob, 
                                    myQueueItem.Identifier.ToString(), 
                                    indexingContainer, 
                                    log);

                log.Info($"Processed blob Result: {JsonConvert.SerializeObject(result)}");

                // REMOVED Entity Framework code for saving results of processing to SQL
            }
            catch (Exception ex)
            {
                log.Error($"Unexpected Error: {ex.Message}{Environment.NewLine}{ex.StackTrace}");
            }
        }

        private async static Task<ProcessingResult> ProcessInputBlobForScanningAsync(
            CloudBlockBlob blob,
            string entityId,
            string indexingContainer,
            TraceWriter log)
        {

            ProcessingResult result = new ProcessingResult();
            string processedBlobName = Path.GetFileNameWithoutExtension(blob.Name);

            try
            {
                blob.FetchAttributes();

                var storageHelper = new StorageHelper();
                var extension = Path.GetExtension(blob.Name).ToLower();
                var builder = new StringBuilder();

                using (var stream = new MemoryStream())
                {
                    await blob.DownloadToStreamAsync(stream);

                    try
                    {
                        // currently only process certain types of documents
                        // if not a processable type of document, we just pass through to indexing location
                        switch (extension)
                        {
                            case ".pdf":

                                var bytes = stream.ToArray();

                                // if there is any text within document add to builder
                                DocumentProcessor.GetTextFromPdf(bytes, out builder);

                                // extract all images within document that are greater than 50x50 pixels
                                List<Stream> images = DocumentProcessor.ExtractImagesFromPDF(bytes, log);

                                if (images.Count > 0)
                                {
                                    int imageCounter = 0;

                                    foreach (Stream img in images)
                                    {
                                        imageCounter++;

                                        try
                                        {
                                            builder.Append(" " + DocumentProcessor.ScanImageToString(img));
                                            log.Info($"OCR completed successfully for pdf image #{imageCounter}");

                                            // Azure Vision service has a cap on images processed per second
                                            // let's slow it down
                                            await Task.Delay(1000);
                                        }
                                        catch (ArgumentException aex)
                                        {
                                            // stream isn't a valid image
                                            log.Warning($"Failed to open image #{imageCounter} of {images.Count} for {blob.Name}. Error:{aex.Message}");
                                            continue;
                                        }
                                        catch (Exception ex)
                                        {
                                            log.Warning($"Failed to OCR scan pfd image #{imageCounter} of {images.Count} for {blob.Name}. Error:{ex.Message}");

                                            // Vision API can throw ClientException, grab inner exception for details
                                            if (ex.InnerException != null && ex.InnerException is ClientException)
                                            {
                                                log.Warning($"InnerException Details: Message={((ClientException)ex.InnerException).Error.Message}");
                                            }
                                        }
                                    }
                                }

                                break;

                            case ".docx":

                                builder.Append(OfficeHelper.GetAllTextFromWordDoc(stream, log));
                                break;

                            case ".xlsx":

                                builder.Append(OfficeHelper.GetAllTextFromExcelDoc(stream, log));
                                break;

                            default:

                                // document is not a proccessable document type.  just send through for indexing
                                result.Status = ProcessingStatus.Success;
                                result.DocumentLocation = await MarkAndSendDocumentAsync(
                                                            entityId,
                                                            blob,
                                                            indexingContainer,
                                                            processedBlobName,
                                                            log);

                                return result;

                        }

                        if (builder.Length == 0)
                            throw new ApplicationException("Text could not be extracted from Document.  Can't create empty document");

                        // we always create a new pdf doc for indexing with all existing text merged with image text
                        using (var textStream = await DocumentProcessor.CreateTextDocumentAsync(builder.ToString()))
                        {
                            log.Info($"Indexable document created successfully!");

                            result.Status = ProcessingStatus.Success;
                            result.DocumentLocation = await MarkAndSendDocumentAsync(
                                                        entityId,
                                                        textStream,
                                                        indexingContainer,
                                                        processedBlobName,
                                                        log);

                            return result;
                        }
                    }
                    catch (ApplicationException aex)
                    {
                        var errorMsg = $"Document failed to get processed.  Passing document along to indexing location";
                        log.Warning(errorMsg);

                        // something went wrong processing document, just send through to get indexed
                        result.Status = ProcessingStatus.Warning;
                        result.Message = $"{errorMsg}. Error:{aex.Message}";
                        result.DocumentLocation = await MarkAndSendDocumentAsync(
                                                    entityId,
                                                    blob,
                                                    indexingContainer,
                                                    processedBlobName,
                                                    log);

                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                result.Status = ProcessingStatus.Failure;
                result.DocumentLocation = null;
                result.Message = $"Failed to process document {blob.Name} due to the following error: {ex.Message}{Environment.NewLine}{ex.StackTrace}";
                log.Error(result.Message);
                return result;
            }
        }

        private async static Task<Uri> MarkAndSendDocumentAsync(
            string entityId,
            CloudBlockBlob blob,
            string indexingContainer,
            string processedBlobName,
            TraceWriter log)
        {
            if (!blob.Metadata.ContainsKey("entityId"))
            {
                blob.Metadata.Add("entityId", entityId);
                await blob.SetMetadataAsync();
            }

            log.Info($"Metadata successfully added to file");

            var storageHelper = new StorageHelper();

            return await storageHelper.CopyBlobToIndexingContainerAsync(
                                        blob,
                                        indexingContainer,
                                        processedBlobName);

        }

        private async static Task<Uri> MarkAndSendDocumentAsync(
            string entityId,
            MemoryStream textStream,
            string indexingContainer,
            string processedBlobName,
            TraceWriter log)
        {
            var metadata = new Dictionary<string, string>
            {
                { "entityId", entityId }
            };

            log.Info($"Metadata successfully added to file");

            var storageHelper = new StorageHelper();

            return await storageHelper.UploadByteArrayAsync(
                                textStream.ToArray(),
                                processedBlobName,
                                indexingContainer,
                                metadata);
        }
    }
}
