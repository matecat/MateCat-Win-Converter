using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.WinConverter.ConversionServer;
using static Translated.MateCAT.WinConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.WinConverter.Converters
{
    [TestClass]
    public class PowerPointConverterTest
    {
        static PowerPointConverter converter;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            converter = new PowerPointConverter();
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            converter.Dispose();
        }

        [TestMethod]
        [DeploymentItem(pptFile, testFilesFolder)]
        [DeploymentItem(pptxFile, testFilesFolder)]
        public void PPTtoPPTXandBack()
        {
            TestConversion(converter, pptFile, FileTypes.ppt, FileTypes.pptx, true);
            TestConversion(converter, pptxFile, FileTypes.pptx, FileTypes.ppt, true);
        }

        [TestMethod]
        [DeploymentItem(ppsFile, testFilesFolder)]
        [DeploymentItem(pptxFile, testFilesFolder)]
        public void PPStoPPTXandBack()
        {
            TestConversion(converter, ppsFile, FileTypes.pps, FileTypes.pptx, true);
            TestConversion(converter, pptxFile, FileTypes.pptx, FileTypes.pps, true);
        }

        [TestMethod]
        [DeploymentItem(potFile, testFilesFolder)]
        [DeploymentItem(pptxFile, testFilesFolder)]
        public void POTtoPPTXandBack()
        {
            TestConversion(converter, potFile, FileTypes.pot, FileTypes.pptx, true);
            TestConversion(converter, pptxFile, FileTypes.pptx, FileTypes.pot, true);
        }

        [TestMethod]
        [DeploymentItem(pptmFile, testFilesFolder)]
        [DeploymentItem(pptxFile, testFilesFolder)]
        public void PPTMtoPPTXandBack()
        {
            TestConversion(converter, pptmFile, FileTypes.pptm, FileTypes.pptx, true);
            TestConversion(converter, pptxFile, FileTypes.pptx, FileTypes.pptm, true);
        }

        [TestMethod]
        [DeploymentItem(potxFile, testFilesFolder)]
        [DeploymentItem(pptxFile, testFilesFolder)]
        public void POTXtoPPTXandBack()
        {
            TestConversion(converter, potxFile, FileTypes.potx, FileTypes.pptx, true);
            TestConversion(converter, pptxFile, FileTypes.pptx, FileTypes.potx, true);
        }

        [TestMethod]
        [DeploymentItem(potmFile, testFilesFolder)]
        [DeploymentItem(pptxFile, testFilesFolder)]
        public void POTMtoPPTXandBack()
        {
            TestConversion(converter, potmFile, FileTypes.potm, FileTypes.pptx, true);
            TestConversion(converter, pptxFile, FileTypes.pptx, FileTypes.potm, true);
        }

    }
}
