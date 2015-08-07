using System;
using System.IO;

namespace LegacyOfficeConverter
{
    public class ConvertersRouter : IConverter, IDisposable
    {
        private PooledConverter<WordConverter> wordConverter;
        private PooledConverter<ExcelConverter> excelConverter;
        private PooledConverter<PowerPointConverter> powerPointConverter;
        private string OCRConsolePath; 

        private bool disposed = false;
        
        public ConvertersRouter(int poolSize, string OCRConsolePath = null)
        {
            wordConverter = new PooledConverter<WordConverter>(poolSize);
            excelConverter = new PooledConverter<ExcelConverter>(poolSize);
            powerPointConverter = new PooledConverter<PowerPointConverter>(poolSize);
            this.OCRConsolePath = OCRConsolePath;
        }

        public void Convert(string inputPath, string outputPath)
        {
            string extension = Path.GetExtension(inputPath).ToLower();
            
            switch (extension)
            {
                case ".doc":
                case ".dot":
                case ".rtf":
                case ".docx":
                    wordConverter.Convert(inputPath, outputPath);
                    break;

                case ".xls":
                case ".xlt":
                    excelConverter.Convert(inputPath, outputPath);
                    break;

                case ".ppt":
                case ".pps":
                case ".pot":
                    powerPointConverter.Convert(inputPath, outputPath);
                    break;

                case ".pdf":
                    new PdfConverter(OCRConsolePath).Convert(inputPath, outputPath);
                    break; 

                case ".jpg":
                case ".tiff":
                case ".png":
                    new OCRConsole(OCRConsolePath).Convert(inputPath, outputPath);
                    break;

                default:
                    // Unsupported exception
                    throw new Exception("FileConverter received an unsupported exception: " + extension + ".");
            }
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
