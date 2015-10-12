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
        [DeploymentItem(xlsxFile, testFilesFolder)]
        public void XLStoXLSXandBack()
        {
            TestConversion(converter, xlsFile, FileTypes.xls, FileTypes.xlsx, true);
            TestConversion(converter, xlsxFile, FileTypes.xlsx, FileTypes.xls, true);
        }

        [TestMethod]
        [DeploymentItem(xltFile, testFilesFolder)]
        [DeploymentItem(xlsxFile, testFilesFolder)]
        public void XLTtoXLSXandBack()
        {
            TestConversion(converter, xltFile, FileTypes.xlt, FileTypes.xlsx, true);
            TestConversion(converter, xlsxFile, FileTypes.xlsx, FileTypes.xlt, true);
        }

        [TestMethod]
        [DeploymentItem(xlsmFile, testFilesFolder)]
        [DeploymentItem(xlsxFile, testFilesFolder)]
        public void XLSMtoXLSXandBack()
        {
            TestConversion(converter, xlsFile, FileTypes.xlsm, FileTypes.xlsx, true);
            TestConversion(converter, xlsxFile, FileTypes.xlsx, FileTypes.xlsm, true);
        }

        [TestMethod]
        [DeploymentItem(xltxFile, testFilesFolder)]
        [DeploymentItem(xlsxFile, testFilesFolder)]
        public void XLTXtoXLSXandBack()
        {
            TestConversion(converter, xlsFile, FileTypes.xltx, FileTypes.xlsx, true);
            TestConversion(converter, xlsxFile, FileTypes.xlsx, FileTypes.xltx, true);
        }

        [TestMethod]
        [DeploymentItem(xltmFile, testFilesFolder)]
        [DeploymentItem(xlsxFile, testFilesFolder)]
        public void XLTMtoXLSXandBack()
        {
            TestConversion(converter, xlsFile, FileTypes.xltm, FileTypes.xlsx, true);
            TestConversion(converter, xlsxFile, FileTypes.xlsx, FileTypes.xltm, true);
        }

        [TestMethod]
        [DeploymentItem(brokenXlsxFile, testFilesFolder)]
        [ExpectedException(typeof(BrokenSourceException))]
        public void BrokenXLSXtoXLS()
        {
            TestConversion(converter, brokenXlsxFile, FileTypes.xlsx, FileTypes.xls, false);
        }

    }
}
