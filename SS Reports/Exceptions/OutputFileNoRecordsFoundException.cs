using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SS_Reports.Exceptions
{
    /// <summary>
    /// There are no records to be subtracted.
    /// </summary>
    class OutputFileNoRecordsFoundException : Exception
    {
        public OutputFileNoRecordsFoundException()
        {
        }
        public OutputFileNoRecordsFoundException(string message) : base(message)
        {
        }
        public OutputFileNoRecordsFoundException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
