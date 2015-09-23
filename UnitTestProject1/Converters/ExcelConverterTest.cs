using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.WinConverter.ConversionServer;
using static Translated.MateCAT.WinConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.WinConverter.Converters
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
            TestConversion(converter, xlsFile, FileTypes.xls, FileTypes.xlsx, true);
        }

        [TestMethod]
        [DeploymentItem(xlsxFile, testFilesFolder)]
        public void XLSXtoXLS()
        {
            TestConversion(converter, xlsxFile, FileTypes.xlsx, FileTypes.xls, true);
        }

    }
}
