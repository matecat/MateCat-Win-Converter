using System.Configuration;
using System.IO;
using System.Net;
using System.Web;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using static Translated.MateCAT.LegacyOfficeConverter.Utils.PdfAnalyzer;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    public class RegularPdfConverter : IConverter
    {
        private static readonly string CloudConverterKey = ConfigurationManager.AppSettings.Get("CloudConvertKey");

        public bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat)
        {
            // Restrict supported conversion to only regular-pdf -> docx, 
            // but don't forget that CloudConvert supports almost anything.
            if (sourceFormat != (int)FileTypes.pdf || targetFormat != (int)FileTypes.docx || IsScannedPdf(sourceFilePath))
            {
                return false;
            }

            // Execute the call
            using (WebClient client = new WebClient())
            {    
                // Send the file to CloudConvert
                client.Headers["Content-Type"] = "binary/octet-stream";
                string address = string.Format("{0}?apikey={1}&input=upload&inputformat={2}&outputformat={3}&file={4}",
                                                             "https://api.cloudconvert.com/convert",
                                                             CloudConverterKey,
                                                             "pdf",
                                                             "docx",
                                                             HttpUtility.UrlEncode(sourceFilePath));
                byte[] response = client.UploadFile(address, sourceFilePath);

                // Save returned converted file
                File.WriteAllBytes(targetFilePath, response);
            }

            // Everything ok, return the success to the caller
            return true;
        }

    }
}
