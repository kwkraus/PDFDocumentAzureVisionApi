using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading.Tasks;

namespace DocProcessing.AzureFunction.Helpers
{
    public class StorageHelper
    {
        private readonly CloudStorageAccount _storageAccount;

        public StorageHelper()
        {
            var connStr = ConfigurationManager.AppSettings["servconstorage"].ToString();
            _storageAccount = CloudStorageAccount.Parse(connStr);
        }

        public async Task<Uri> UploadByteArrayAsync(
            byte[] fileContents, 
            string fileName, 
            string indexingContainer,
            Dictionary<string, string> metadata = null)
        {
            CloudBlobClient blobClient = _storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference(indexingContainer);

            CloudBlockBlob blockBlob = container.GetBlockBlobReference(fileName);

            // Create the container if it doesn't already exist.
            await container.CreateIfNotExistsAsync();

            await blockBlob.UploadFromByteArrayAsync(fileContents, 0, fileContents.Length);

            if (metadata != null)
            {
                foreach (var item in metadata)
                {
                    // add filing id guid to blob metadata.  this will be used by Azure Search Indexer to map to index document
                    blockBlob.Metadata.Add(item.Key, item.Value);
                }
                await blockBlob.SetMetadataAsync();
            }

            return blockBlob.Uri;
        }

        public async Task<Uri> CopyBlobToIndexingContainerAsync(
            CloudBlockBlob sourceBlob, 
            string indexingContainer,
            string blobName = null)
        {
            CloudBlobClient cloudBlobClient = _storageAccount.CreateCloudBlobClient();
            CloudBlobContainer targetContainer = cloudBlobClient.GetContainerReference(indexingContainer);
            await targetContainer.CreateIfNotExistsAsync();

            if (string.IsNullOrEmpty(blobName))
                blobName = sourceBlob.Name;

            CloudBlockBlob targetBlob = targetContainer.GetBlockBlobReference(blobName);
            await targetBlob.StartCopyAsync(sourceBlob);
            return targetBlob.Uri;
        }
    }
}
