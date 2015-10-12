using System;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class BrokenSourceException : Exception
    {
        public BrokenSourceException() : base() {}
        public BrokenSourceException(string message) : base(message) { }
        public BrokenSourceException(string message, Exception innerException) : base(message, innerException) { }
    }
}
