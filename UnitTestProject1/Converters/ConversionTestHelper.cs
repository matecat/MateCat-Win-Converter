using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    class ConversionTestHelper
    {
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

        public static void TestConversionDone(IConverter converter, string sourceFileName, FileTypes sourceFormat, FileTypes targetFormat)
        {
            string sourceFile = Path.GetFullPath(sourceFileName);
            string targetFile = Path.GetTempFileName();

            bool converted = converter.Convert(sourceFile, (int)sourceFormat, targetFile, (int)targetFormat);

            // Check that converter returned true
            Assert.AreEqual(true, converted);
            // Check that converter created the target file and it's not empty
            Assert.AreNotEqual(0, new FileInfo(targetFile).Length);

            // Delete the temp file created
            if (File.Exists(targetFile)) File.Delete(targetFile);
        }

        public static void TestConversionSkipped(IConverter converter, string sourceFileName, FileTypes sourceFormat, FileTypes targetFormat)
        {
            string sourceFile = Path.GetFullPath(sourceFileName);
            string targetFile = Path.GetTempFileName();

            bool converted = converter.Convert(sourceFile, (int)sourceFormat, targetFile, (int)targetFormat);

            // Check that converter returned false
            Assert.AreEqual(false, converted);

            // Delete the temp file created
            if (File.Exists(targetFile)) File.Delete(targetFile);
        }

    }
}
