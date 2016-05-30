using System;

namespace SS_Reports.Exceptions
{
    /// <summary>
    /// The source file does not match the file selected.
    /// </summary>
    class SourceFileNotMatchingSelectedFileException : Exception
    {
        public SourceFileNotMatchingSelectedFileException()
        {
        }
        public SourceFileNotMatchingSelectedFileException(string message) : base(message)
        {
        }
        public SourceFileNotMatchingSelectedFileException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
