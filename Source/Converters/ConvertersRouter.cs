using System;
using System.Collections.Generic;
using log4net;
using static System.Reflection.MethodBase;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class ConvertersRouter : IConverter, IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        private readonly IList<IConverter> converters;

        private bool disposed = false;
        
        public ConvertersRouter(int poolSize)
        {
            // Create the list with all supported converters
            List<IConverter> converters = new List<IConverter>();
            converters.Add(new PooledConverter<WordConverter>(poolSize));
            converters.Add(new PooledConverter<ExcelConverter>(poolSize));
            converters.Add(new PooledConverter<PowerPointConverter>(poolSize));
            converters.Add(new OcrConverter());
            converters.Add(new RegularPdfConverter());

            this.converters = converters.AsReadOnly();
        }

        public bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat)
        {
            // Try to delegate the conversion to each converter in the list,
            // until one of them performs it. If there aren't converters that
            // support the required conversion, return false.
            bool converted = false;
            foreach (IConverter converter in converters)
            {
                converted = converter.Convert(sourceFilePath, sourceFormat, targetFilePath, targetFormat);
                if (converted) break;
            }
            return converted;
        }

        
        /*
         * Pay great attention to the dispose/destruction functions, it's
         * very important to release the used Office objects properly.
         */
        public void Dispose()
        { 
            Dispose(true);
            GC.SuppressFinalize(this);           
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return; 

            if (disposing) {
                foreach (IConverter converter in converters)
                {
                    if (converter is IDisposable)
                    {
                        ((IDisposable)converter).Dispose();
                    }
                }
                converters.Clear();
            }
            disposed = true;
        }

        ~ConvertersRouter()
        {
            Dispose(false);
        }
    }
}
