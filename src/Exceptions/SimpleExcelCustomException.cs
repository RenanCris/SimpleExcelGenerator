using System;
using System.Runtime.Serialization;

namespace SimpleExcelGenerator.Exceptions
{
    [Serializable]
    public class SimpleExcelCustomException : Exception
    {
        public SimpleExcelCustomException()
        {
        }

        public SimpleExcelCustomException(string message) : base(message)
        {
        }

        public SimpleExcelCustomException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected SimpleExcelCustomException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
