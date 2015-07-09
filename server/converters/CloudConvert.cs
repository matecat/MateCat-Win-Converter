using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System;
using System.Collections.Generic;
using System.Net;
using System.Web;
using System.Web.Script.Serialization;
using System.Text;

namespace LegacyOfficeConverter
{

    class CloudConvert
    {
        private static readonly string CloudConverterKey = "zSOdJP2QupwbNgSCO7cq2sWlhIKqiU28-tRDnUF1-T40Q8lYOWSp_u4l2Sms-oJzdYlM3LsW6gx0FO88KXAQAw";
        private static readonly string BaseUrl = "https://api.cloudconvert.com/convert";

        // TODO: add exception handling
        public void Convert(string inputPath, string outputPath)
        {
            // Compute extensions
            string inputFormat = Path.GetExtension(inputPath).Substring(1);
            string outputFormat = Path.GetExtension(outputPath).Substring(1);

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
