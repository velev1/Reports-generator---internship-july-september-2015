using System;

namespace SS_Reports.Exceptions
{
    /// <summary>
    /// Indicates that the output file is not matching the predefined format.
    /// </summary>
    class OutputFileNotCorrectException : Exception
    {
        public OutputFileNotCorrectException()
        {
        }
        public OutputFileNotCorrectException(string message) : base(message)
        {
        }
        public OutputFileNotCorrectException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
