using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SS_Reports.Exceptions
{
    /// <summary>
    /// Indicates that the output file is full.
    /// </summary>
    public class OutputFileIsFullException : Exception
    {
        public OutputFileIsFullException()
        {
        }
        public OutputFileIsFullException(string message) : base(message)
        {
        }
        public OutputFileIsFullException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
