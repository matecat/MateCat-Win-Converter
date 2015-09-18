using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using static Translated.MateCAT.LegacyOfficeConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    [TestClass]
    public class ExcelConverterTest
    {
        static ExcelConverter converter;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            converter = new ExcelConverter();
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            converter.Dispose();
        }

        [TestMethod]
        [DeploymentItem(xlsFile, testFilesFolder)]
        public void XLStoXLSX()
        {
            TestConversionDone(converter, xlsFile, FileTypes.xls, FileTypes.xlsx);
        }

        [TestMethod]
        [DeploymentItem(xlsxFile, testFilesFolder)]
        public void XLSXtoXLS()
        {
            TestConversionDone(converter, xlsxFile, FileTypes.xlsx, FileTypes.xls);
        }

    }
}
