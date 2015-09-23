using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.WinConverter.ConversionServer;
using static Translated.MateCAT.WinConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.WinConverter.Converters
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
            TestConversion(converter, docFile, FileTypes.doc, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(docxFile, testFilesFolder)]
        public void DOCXtoDOC()
        {
            TestConversion(converter, docxFile, FileTypes.docx, FileTypes.doc, true);
        }

        [TestMethod]
        [DeploymentItem(dotFile, testFilesFolder)]
        public void DOTtoDOCX()
        {
            TestConversion(converter, dotFile, FileTypes.dot, FileTypes.docx, true);
        }

        [TestMethod]
        [DeploymentItem(rtfFile, testFilesFolder)]
        public void RTFtoDOCX()
        {
            TestConversion(converter, rtfFile, FileTypes.rtf, FileTypes.docx, true);
        }

    }
}
