using Newtonsoft.Json;
using System;
using System.IO;

namespace DocProcessing.AzureFunction.Models
{
    public class BlobInfo
    {
        //requires parameterless constructor for deserialization in webjob
        public BlobInfo() { }

        public BlobInfo(Uri blobLocation)
        {
            BlobUri = blobLocation;

            //validate the Uri by trying to parse the filingid
            try
            {
                Guid id = Identifier;
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Blob Name is not a valid Guid {blobLocation.ToString()}", ex);
            }
        }

        [JsonProperty]
        public Uri BlobUri { get; set; }

        public string BlobName
        {
            get
            {
                return BlobUri.Segments[BlobUri.Segments.Length - 1];
            }
        }

        public Guid Identifier
        {
            get
            {
                return new Guid(Path.GetFileNameWithoutExtension(BlobName));
            }
        }
    }
}
