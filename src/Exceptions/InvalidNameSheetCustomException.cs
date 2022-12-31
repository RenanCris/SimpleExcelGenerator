using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace SimpleExcelGenerator.Exceptions
{
    [Serializable]
    public class InvalidNameSheetCustomException : Exception
    {
        public InvalidNameSheetCustomException()
        {
        }

        public InvalidNameSheetCustomException(string message) : base(message)
        {
        }

        public InvalidNameSheetCustomException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected InvalidNameSheetCustomException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
