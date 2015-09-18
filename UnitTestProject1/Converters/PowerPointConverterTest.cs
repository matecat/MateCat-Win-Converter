using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using static Translated.MateCAT.LegacyOfficeConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
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
        public void PPTtoPPTX()
        {
            TestConversionDone(converter, pptFile, FileTypes.ppt, FileTypes.pptx);
        }

        [TestMethod]
        [DeploymentItem(pptxFile, testFilesFolder)]
        public void PPTXtoPPT()
        {
            TestConversionDone(converter, pptxFile, FileTypes.pptx, FileTypes.ppt);
        }

    }
}
