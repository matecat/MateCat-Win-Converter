using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using Translated.MateCAT.LegacyOfficeConverter.Utils;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    class ConversionTestHelper
    {
        private static readonly object fileSystemLock = new object();

        public const string testFilesFolder = "TestFiles";
        public const string pptFile = testFilesFolder + "\\source.ppt";
        public const string pptxFile = testFilesFolder + "\\source.pptx";
        public const string docFile = testFilesFolder + "\\source.doc";
        public const string docxFile = testFilesFolder + "\\source.docx";
        public const string dotFile = testFilesFolder + "\\source.dot";
        public const string xlsFile = testFilesFolder + "\\source.xls";
        public const string xlsxFile = testFilesFolder + "\\source.xlsx";
        public const string rtfFile = testFilesFolder + "\\source.rtf";
        public const string jpgFile = testFilesFolder + "\\source.jpg";
        public const string pdfFile = testFilesFolder + "\\source.pdf";
        public const string scannedPdfFile = testFilesFolder + "\\scanned_source.pdf";

        public static void TestConversion(IConverter converter, string sourceFileName, FileTypes sourceFormat, FileTypes targetFormat, bool assertSuccess)
        {
            string sourceFile = Path.GetFullPath(sourceFileName);

            // Get a temporary target file
            TempFolder tempFolder = new TempFolder();
            string targetFile = tempFolder.getFilePath("target." + targetFormat);

            // Do the conversion
            bool converted = converter.Convert(sourceFile, (int)sourceFormat, targetFile, (int)targetFormat);

            if (assertSuccess)
            {
                // Check that converter returned true
                Assert.AreEqual(true, converted);
                // Check that converter created the target file and it's not empty
                Assert.AreNotEqual(0, new FileInfo(targetFile).Length);
            }
            else
            {
                // Check that converter returned false
                Assert.AreEqual(false, converted);
            }

            // Delete the temp file created
            if (File.Exists(targetFile)) File.Delete(targetFile);
        }

        public static void TestHighConcurrencyConversion(IConverter converter, string sourceFileName, FileTypes sourceFormat, FileTypes targetFormat, int concurrentConversions)
        {
            // Try to convert always the same file
            string sourceFile = Path.GetFullPath(sourceFileName);

            // Count the succeded conversion in this var
            int successes = 0;

            // Create the thread pool
            List<Thread> threads = new List<Thread>();
            for (int i = 0; i < concurrentConversions; i++)
            {
                threads.Add(new Thread(delegate () {
                    // This code will run many times in the same moment on many threads,
                    // so be very careful to not make two Office instances write the same
                    // file, or they will raise exceptions. The most problems I had 
                    // happened exactly because of this. Always make Office instance write
                    // to a clean filepath that points to a still non-existent file.
                    // Don't try anything else, 90% it will raise errors.

                    // The safest way to create a clean lonely filepath is using the 
                    // TempFolder class
                    TempFolder tempFolder = new TempFolder();
                    string targetFile = tempFolder.getFilePath("target." + targetFormat);

                    // Ok, let's do the real conversion
                    try
                    {
                        bool converted = converter.Convert(sourceFile, (int)sourceFormat, targetFile, (int)targetFormat);
                        if (converted) Interlocked.Increment(ref successes);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Conversion failed:\n" + e + "\n");
                    }
                    finally
                    {
                        // Delete the temp target folder
                        Directory.Delete(tempFolder.ToString(), true);
                    }
                }));
            }

            // Launch all the threads in the same time
            foreach (var thread in threads)
            {
                thread.Start();
            }

            // Wait for all threads completion
            foreach (var thread in threads)
            {
                thread.Join();
            }

            // Check test final result
            Assert.AreEqual(concurrentConversions, successes);
        }

    }
}
