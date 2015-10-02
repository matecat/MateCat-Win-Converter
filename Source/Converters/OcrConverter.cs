using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Translated.MateCAT.WinConverter.ConversionServer;
using static Translated.MateCAT.WinConverter.Utils.PdfAnalyzer;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class OcrConverter : IConverter
    {
        private static readonly string ocrConsolePath = ConfigurationManager.AppSettings.Get("OCRConsolePath");
        private static readonly bool isInstalled;

        private static readonly int[] validSourceExtensions = { (int)FileTypes.jpg, (int)FileTypes.png, (int)FileTypes.tiff, (int)FileTypes.pdf };

        static OcrConverter()
        {
            isInstalled = File.Exists(ocrConsolePath);
            if (!isInstalled)
            {
                Console.WriteLine("WARNING: can't find OCR Console executable in path \"" + ocrConsolePath + "\": OCR conversions disabled");
            }
        }

        public bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat)
        {
            // Check if the required conversion is supported
            if (!validSourceExtensions.Contains(sourceFormat) || targetFormat != (int)FileTypes.docx)
            {
                return false;
            }
            // If we have a PDF and it's not scanned, return false
            if (sourceFormat == (int)FileTypes.pdf && !IsScannedPdf(sourceFilePath))
            {
                return false;
            }

            // If OCR Console is not installed, skip the conversion
            if (!isInstalled)
            {
                Console.WriteLine("Skipped OCR Conversion because OCR Console is missing");
                return false;
            }

            // Setup process
            var process = new Process();
            process.StartInfo.Arguments = "\"" + sourceFilePath + "\" \"" + targetFilePath + "\"";
            process.StartInfo.FileName = ocrConsolePath;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.RedirectStandardOutput = true;

            // Execute it
            process.Start();
            process.WaitForExit();

            // Check errors
            if (process.ExitCode != 0)
            {
                throw new Exception("OCR Console returned exit code " + process.ExitCode);
            }

            // Everything ok, return the success to the caller        
            return true;
        }

        public static bool IsInstalled()
        {
            return isInstalled;
        }

    }
}
