using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using static Translated.MateCAT.LegacyOfficeConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    [TestClass]
    public class CloudConvertTest
    {
        static CloudConvert converter;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            converter = new CloudConvert();
        }

        [TestMethod]
        [DeploymentItem(pdfFile, testFilesFolder)]
        public void RegularPDFtoDOCX()
        {
            TestConversion(converter, pdfFile, FileTypes.pdf, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(scannedPdfFile, testFilesFolder)]
        public void CloudConvertConverterSkipsScannedPDF()
        {
            TestConversion(converter, scannedPdfFile, FileTypes.pdf, FileTypes.docx, false);
        }

    }
}
