using System;

namespace DocProcessing.AzureFunction
{
    public class ProcessingResult
    {
        public ProcessingStatus Status { get; set; }
        public string Message { get; set; }
        public Uri DocumentLocation { get; set; }
    }

    public enum ProcessingStatus
    {
        Success,
        Failure,
        Warning
    }
}