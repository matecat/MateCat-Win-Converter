using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    class PdfConverter : IConverter
    {

        private OcrConverter ocrConsole;
        private RegularPdfConverter cloudConvert;

        public PdfConverter(string OCRConsolePath = null)
        {
            ocrConsole = new OcrConverter();
            cloudConvert = new RegularPdfConverter();
        }


        /// <summary>
        /// Convert both normal and scanned PDFs
        /// </summary>
        public bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat)
        {
            // Check if the required conversion is supported
            if (sourceFormat != (int)FileTypes.pdf || targetFormat != (int)FileTypes.docx)
            {
                return false;
            }

            // Decide which converter is better
            IConverter pdfConverter = (IsScannedPDF(sourceFilePath) ? (IConverter)ocrConsole : (IConverter)cloudConvert);

            // Perform the conversion
            return pdfConverter.Convert(sourceFilePath, sourceFormat, targetFilePath, targetFormat);
        }

        public static bool IsScannedPDF(string pdfFilePath)
        {
            // Start analyzing the PDF
            PdfReader reader = new PdfReader(pdfFilePath);
            PdfDictionary resources;

            // Go through all the pages
            for (int p = 1; p <= reader.NumberOfPages; p++)
            {
                // Find the embedded resources
                PdfDictionary dic = reader.GetPageN(p);
                resources = dic.GetAsDict(PdfName.RESOURCES);
                if (resources != null)
                {
                    // If we have any embedded font, it's not scanned
                    if (resources.GetAsDict(PdfName.FONT) != null)
                        return false;
                }
            }
            return true;
        }

    }
}
