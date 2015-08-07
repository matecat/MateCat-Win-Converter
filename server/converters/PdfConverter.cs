using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    class PdfConverter : IConverter
    {

        private string OCRConsolePath;

        public PdfConverter(string OCRConsolePath = null)
        {
            this.OCRConsolePath = OCRConsolePath;
        }


        /// <summary>
        /// Convert both normal and scanned PDFs
        /// </summary>
        /// <param name="inputPath">Input path</param>
        /// <param name="outputPath">Output path</param>
        public void Convert(string inputPath, string outputPath)
        {

            // Check that we have are working with a PDF file and that it exists
            if (Path.GetExtension(inputPath).ToLower() != ".pdf" || !File.Exists(inputPath))
                throw new ArgumentException("The given file is not a PDF");
            if (Path.GetExtension(outputPath).ToLower() != ".docx")
                throw new ArgumentException("PDF files can just be converted to DOCX");

            // If its scanned, execute an OCR recognition
            if (new PdfAnalyzer(inputPath).IsScanned())
            {
               if (this.OCRConsolePath == null)
                    throw new Exception("A OCR console path has not been specified");
               new OCRConsole(OCRConsolePath).Convert(inputPath, outputPath);
            }

            // If not, convert it through cloudconvert
            else
            {
                new CloudConvert().Convert(inputPath, outputPath); 
            }

            
        }
    }
}
