using iTextSharp.text.pdf;

namespace Translated.MateCAT.WinConverter.Utils
{
    public class PdfAnalyzer
    {

        public static bool IsScannedPdf(string pdfFilePath)
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
