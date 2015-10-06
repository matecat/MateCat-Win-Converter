using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.WinConverter.ConversionServer;
using static Translated.MateCAT.WinConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.WinConverter.Converters
{
    [TestClass]
    public class OcrConverterTest
    {
        static OcrConverter converter;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            converter = new OcrConverter();
        }

        [TestMethod]
        [DeploymentItem(scannedPdfFile, testFilesFolder)]
        public void ScannedPDFtoDOCX()
        {
            TestConversion(converter, scannedPdfFile, FileTypes.pdf, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(jpegFile, testFilesFolder)]
        public void JPEGtoDOCX()
        {
            TestConversion(converter, jpegFile, FileTypes.jpeg, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(scannedPdfFile, testFilesFolder)]
        public void OcrConverterSkipsRegularPDF()
        {
            TestConversion(converter, scannedPdfFile, FileTypes.pdf, FileTypes.docx, false);
        }

    }
}
