using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXmlCrypto
{
    public class InvalidPasswordException : ApplicationException
    {
        public InvalidPasswordException() : base() { }
        public InvalidPasswordException(String message) : base(message) { }
        public InvalidPasswordException(String message, Exception innerException)
            : base(message, innerException) { }
    }
}
