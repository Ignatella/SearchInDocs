using System;

namespace WordLibrary
{
    public class FileIsProcessedEventArgs : EventArgs
    {
        string FileName { get; }
        public FileIsProcessedEventArgs(string fileName)
        {
            FileName = fileName;
        }
    }
}
