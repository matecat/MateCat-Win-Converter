using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using static Translated.MateCAT.LegacyOfficeConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    [TestClass]
    public class WordConverterTest
    {
        static WordConverter converter;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            converter = new WordConverter();
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            converter.Dispose();
        }

        [TestMethod]
        [DeploymentItem(docFile, testFilesFolder)]
        public void DOCtoDOCX()
        {
            TestConversionDone(converter, docFile, FileTypes.doc, FileTypes.docx);
        }

        [TestMethod]
        [DeploymentItem(docxFile, testFilesFolder)]
        public void DOCXtoDOC()
        {
            TestConversionDone(converter, docxFile, FileTypes.docx, FileTypes.doc);
        }

        [TestMethod]
        [DeploymentItem(dotFile, testFilesFolder)]
        public void DOTtoDOCX()
        {
            TestConversionDone(converter, dotFile, FileTypes.dot, FileTypes.docx);
        }

        [TestMethod]
        [DeploymentItem(rtfFile, testFilesFolder)]
        public void RTFtoDOCX()
        {
            TestConversionDone(converter, rtfFile, FileTypes.rtf, FileTypes.docx);
        }

    }
}
