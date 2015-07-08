using System;
using System.IO;

namespace LegacyOfficeConverter
{
    public class ConvertersRouter : IConverter
    {
        private PooledConverter<WordConverter> wordConverter;
        private PooledConverter<ExcelConverter> excelConverter;
        private PooledConverter<PowerPointConverter> powerPointConverter;
        private OCRConsole ocrConverter;


        private bool disposed = false;
        
        public ConvertersRouter(int poolSize, string OCRConsolePath = null)
        {
            wordConverter = new PooledConverter<WordConverter>(poolSize);
            excelConverter = new PooledConverter<ExcelConverter>(poolSize);
            powerPointConverter = new PooledConverter<PowerPointConverter>(poolSize);

            // Store the path of the OCRConsole
            ocrConverter = new OCRConsole(OCRConsolePath);
        }

        public string Convert(string path)
        {
            string extension = Path.GetExtension(path).ToLower();
            string convertedPath = null;

            switch (extension)
            {
                case ".doc":
                case ".dot":
                case ".rtf":
                    convertedPath = wordConverter.Convert(path);
                    break;

                case ".xls":
                case ".xlt":
                    convertedPath = excelConverter.Convert(path);
                    break;

                case ".ppt":
                case ".pps":
                case ".pot":
                    convertedPath = powerPointConverter.Convert(path);
                    break;

                case ".pdf":
                case ".jpg":
                case ".tiff":
                case ".png":
                    convertedPath = convertOCR(path);
                    break;

                default:
                    // Unsupported exception
                    throw new Exception("FileConverter received an unsupported exception: " + extension + ".");
            }

            return convertedPath;
        }


        /**
         * Convert OCR file, return its converted path
         */
        private string convertOCR(string path)
        {
            return ocrConverter.OCRrecognition(path);
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
                wordConverter.Dispose();
                excelConverter.Dispose();
                powerPointConverter.Dispose();
            }

            disposed = true;
        }

        ~ConvertersRouter()
        {
            Dispose(false);
        }
    }
}
