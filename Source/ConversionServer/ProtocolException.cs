using System;

namespace Translated.MateCAT.WinConverter.ConversionServer
{
    public class ProtocolException : Exception
    {
        public readonly StatusCodes statusCode;

        public ProtocolException(StatusCodes statusCode, Exception innerException)
            : base("Exception " + (int)statusCode + ": " + statusCode, innerException)
        {
            this.statusCode = statusCode;
        }

        public ProtocolException(StatusCodes errorCode) 
            : this(errorCode, null)
        { }

    }
}
