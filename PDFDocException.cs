using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Adhoc_Adobe_Library_3_5
{
    class PDFDocException : Exception
    {
        public PDFDocException() : base() { }

        public PDFDocException(string messageValue) : base(messageValue) { }
    }
}
