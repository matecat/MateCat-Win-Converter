using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    public class ProtocolException : Exception
    {
        public readonly Errors ErrorCode;

        public ProtocolException(Errors errorCode, Exception innerException)
            : base("Exception " + (int)errorCode + ": " + errorCode, innerException)
        {
            this.ErrorCode = errorCode;
        }

        public ProtocolException(Errors errorCode) 
            : this(errorCode, null)
        { }

    }
}
