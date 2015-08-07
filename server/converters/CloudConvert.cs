using System;
using System.IO;
using System.Net;
using System.Web;

namespace LegacyOfficeConverter
{

    class CloudConvert : IConverter
    {
        private static readonly string CloudConverterKey = "zSOdJP2QupwbNgSCO7cq2sWlhIKqiU28-tRDnUF1-T40Q8lYOWSp_u4l2Sms-oJzdYlM3LsW6gx0FO88KXAQAw";
        private static readonly string BaseUrl = "https://api.cloudconvert.com/convert";

        // TODO: add exception handling
        public void Convert(string inputPath, string outputPath)
        {
            // Compute extensions
            string inputFormat = Path.GetExtension(inputPath).Replace(".", "");
            string outputFormat = Path.GetExtension(outputPath).Replace(".", "");

            // Execute the call
            using (WebClient client = new WebClient())
            {    

               // Update the file and get the response
               client.Headers["Content-Type"] = "binary/octet-stream";
               byte[] response = client.UploadFile(string.Format("{0}?apikey={1}&input=upload&inputformat={2}&outputformat={3}&file={4}",
                                                             BaseUrl,
                                                             CloudConverterKey,
                                                             inputFormat,
                                                             outputFormat,
                                                             HttpUtility.UrlEncode(inputPath)),
                                               inputPath);

                // Compute the out path and save it
                File.WriteAllBytes(outputPath, response);
            }
        }

    }
}
