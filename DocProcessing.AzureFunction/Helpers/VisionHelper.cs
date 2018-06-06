using System.Text;
using System.IO;
using Microsoft.ProjectOxford.Vision;
using Microsoft.ProjectOxford.Vision.Contract;

namespace DocProcessing.AzureFunction.Helpers
{

    /// <summary>
    /// The class is used to access vision APIs.
    /// </summary>
    public class VisionHelper
    {
        /// <summary>
        /// The vision service client.
        /// </summary>
        private readonly IVisionServiceClient visionClient;

        /// <summary>
        /// Initializes a new instand of the <see cref="VisionHelper"/> class.
        /// </summary>
        /// <param name="subscriptionKey">The subscription key.</param>
        public VisionHelper(string subscriptionKey)
        {
            this.visionClient = new VisionServiceClient(subscriptionKey);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="VisionHelper"/> class.
        /// </summary>
        /// <param name="subscriptionKey">The subscription key.</param>
        /// <param name="rootApi">Root Api Url.</param>
        public VisionHelper(string subscriptionKey, string rootApi)
        {
            this.visionClient = new VisionServiceClient(subscriptionKey, rootApi);
        }

        /// <summary>
        /// Recognize text from given image.
        /// </summary>
        /// <param name="imagePathOrUrl">The image path or url.</param>
        /// <param name="detectOrientation">if set to <c>true</c> [detect orientation].</param>
        /// <param name="languageCode">The language code.</param>
        public OcrResults RecognizeText(Stream imageStream, bool detectOrientation = true, string languageCode = LanguageCodes.AutoDetect)
        {
            OcrResults ocrResult = null;
            string resultStr = string.Empty;

            return ocrResult = this.visionClient.RecognizeTextAsync(imageStream, languageCode, detectOrientation).Result;
        }


        /// <summary>
        /// Retrieve text from the given OCR results object.
        /// </summary>
        /// <param name="results">The OCR results.</param>
        /// <returns>Return the text.</returns>
        public string GetRetrieveText(OcrResults results)
        {
            StringBuilder stringBuilder = new StringBuilder();

            if (results != null && results.Regions != null)
            {
                stringBuilder.AppendLine();
                foreach (var item in results.Regions)
                {
                    foreach (var line in item.Lines)
                    {
                        foreach (var word in line.Words)
                        {
                            stringBuilder.Append(word.Text);
                            stringBuilder.Append(" ");
                        }

                        stringBuilder.AppendLine();
                    }

                    stringBuilder.AppendLine();
                }
            }

            return stringBuilder.ToString();
        }
    }
}

