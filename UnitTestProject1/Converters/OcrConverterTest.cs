using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using static Translated.MateCAT.LegacyOfficeConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
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
            TestConversionDone(converter, scannedPdfFile, FileTypes.pdf, FileTypes.docx);
        }

        [TestMethod]
        [DeploymentItem(jpgFile, testFilesFolder)]
        public void JPGtoDOCX()
        {
            TestConversionDone(converter, jpgFile, FileTypes.jpg, FileTypes.docx);
        }

        [TestMethod]
        [DeploymentItem(scannedPdfFile, testFilesFolder)]
        public void OcrConverterSkipsRegularPDF()
        {
            TestConversionSkipped(converter, scannedPdfFile, FileTypes.pdf, FileTypes.docx);
        }

    }
}
