using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using Translated.MateCAT.WinConverter.ConversionServer;
using Translated.MateCAT.WinConverter.Utils;
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
        [DeploymentItem(docxFile, testFilesFolder)]
        public void DOCtoDOCXandBack()
        {
            TestConversion(converter, docFile, FileTypes.doc, FileTypes.docx, true);
            TestConversion(converter, docxFile, FileTypes.docx, FileTypes.doc, true);
        }

        [TestMethod]
        [DeploymentItem(dotFile, testFilesFolder)]
        [DeploymentItem(docxFile, testFilesFolder)]
        public void DOTtoDOCXandBack()
        {
            TestConversion(converter, dotFile, FileTypes.dot, FileTypes.docx, true);
            TestConversion(converter, docxFile, FileTypes.docx, FileTypes.dot, true);
        }

        [TestMethod]
        [DeploymentItem(docmFile, testFilesFolder)]
        [DeploymentItem(docxFile, testFilesFolder)]
        public void DOCMtoDOCXandBack()
        {
            TestConversion(converter, docmFile, FileTypes.docm, FileTypes.docx, true);
            TestConversion(converter, docxFile, FileTypes.docx, FileTypes.docm, true);
        }

        [TestMethod]
        [DeploymentItem(dotxFile, testFilesFolder)]
        [DeploymentItem(docxFile, testFilesFolder)]
        public void DOTXtoDOCXandBack()
        {
            TestConversion(converter, dotxFile, FileTypes.dotx, FileTypes.docx, true);
            TestConversion(converter, docxFile, FileTypes.docx, FileTypes.dotx, true);
        }

        [TestMethod]
        [DeploymentItem(dotmFile, testFilesFolder)]
        [DeploymentItem(docxFile, testFilesFolder)]
        public void DOTMtoDOCXandBack()
        {
            TestConversion(converter, dotmFile, FileTypes.dotm, FileTypes.docx, true);
            TestConversion(converter, docxFile, FileTypes.docx, FileTypes.dotm, true);
        }

        [TestMethod]
        [DeploymentItem(rtfFile, testFilesFolder)]
        [DeploymentItem(docxFile, testFilesFolder)]
        public void RTFtoDOCXandBack()
        {
            TestConversion(converter, rtfFile, FileTypes.rtf, FileTypes.docx, true);
            TestConversion(converter, docxFile, FileTypes.docx, FileTypes.rtf, true);
        }

        [TestMethod]
        [DeploymentItem(advancedDocxFile, testFilesFolder)]
        public void AdvancedDOCXtoDOC()
        {
            TestConversion(converter, advancedDocxFile, FileTypes.docx, FileTypes.doc, true);
        }

    }
}
