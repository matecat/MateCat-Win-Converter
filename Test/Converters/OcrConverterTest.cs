using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.WinConverter.ConversionServer;
using static Translated.MateCAT.WinConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.WinConverter.Converters
{
    [TestClass]
    public class OcrConverterTest
    {
        static OcrConverter converter;
        static string missingOcrConsole = "OCR Console needed for this test, but not found.";

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            converter = new OcrConverter();
        }

        [TestMethod]
        [DeploymentItem(scannedPdfFile, testFilesFolder)]
        public void OcrConverterSkipsRegularPDF()
        {
            TestConversion(converter, scannedPdfFile, FileTypes.pdf, FileTypes.docx, false);
        }

        [TestMethod]
        [DeploymentItem(scannedPdfFile, testFilesFolder)]
        public void ScannedPDFtoDOCX()
        {
            if (!OcrConverter.IsInstalled()) Assert.Inconclusive(missingOcrConsole);
            TestConversion(converter, scannedPdfFile, FileTypes.pdf, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(bmpFile, testFilesFolder)]
        public void BMPtoDOCX()
        {
            if (!OcrConverter.IsInstalled()) Assert.Inconclusive(missingOcrConsole);
            TestConversion(converter, bmpFile, FileTypes.bmp, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(jpegFile, testFilesFolder)]
        public void JPEGtoDOCX()
        {
            if (!OcrConverter.IsInstalled()) Assert.Inconclusive(missingOcrConsole);
            TestConversion(converter, jpegFile, FileTypes.jpeg, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(pngFile, testFilesFolder)]
        public void PNGtoDOCX()
        {
            if (!OcrConverter.IsInstalled()) Assert.Inconclusive(missingOcrConsole);
            TestConversion(converter, pngFile, FileTypes.png, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(tiffFile, testFilesFolder)]
        public void TIFFtoDOCX()
        {
            if (!OcrConverter.IsInstalled()) Assert.Inconclusive(missingOcrConsole);
            TestConversion(converter, tiffFile, FileTypes.tiff, FileTypes.docx, true);
        }

    }
}
