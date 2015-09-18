using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Translated.MateCAT.LegacyOfficeConverter.ConversionServer
{
    public class ProtocolException : Exception
    {
        public readonly StatusCodes ErrorCode;

        public ProtocolException(StatusCodes errorCode, Exception innerException)
            : base("Exception " + (int)errorCode + ": " + errorCode, innerException)
        {
            this.ErrorCode = errorCode;
        }

        public ProtocolException(StatusCodes errorCode) 
            : this(errorCode, null)
        { }

    }
}
