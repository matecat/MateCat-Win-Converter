using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace LegacyOfficeConverter
{
    class PdfAnalyzer
    {

        private string path;

        public PdfAnalyzer(string path)
        {
            // Check that we are processing a PDF
            if (Path.GetExtension(path).ToLower() != ".pdf" || !File.Exists(path))
                throw new ArgumentException("The given file is not a PDF");
            this.path = path;
        }

        public bool IsScanned()
        {
            // Start analyzing the PDF
            PdfReader reader = new PdfReader(path);
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
