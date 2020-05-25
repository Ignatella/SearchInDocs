using System;

namespace WordLibrary
{
    public class ErrorOccuredEventArgs : EventArgs
    {
        string FileName { get; }

        Exception Ex { get; }

        public ErrorOccuredEventArgs(string fileName, Exception ex)
        {
            FileName = fileName;
            Ex = ex;
        }
    }
}
