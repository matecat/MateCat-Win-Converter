using Microsoft.VisualStudio.TestTools.UnitTesting;
using Translated.MateCAT.WinConverter.ConversionServer;
using static Translated.MateCAT.WinConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.WinConverter.Converters
{
    [TestClass]
    public class PooledConverterTest
    {
        // The size of the converters' pool
        const int convertersPoolSize = 10;
        // The number of concurrent threads that will launch a conversion 
        // on the same PooledConverter
        const int concurrentConversions = 50;

        [TestMethod]
        [DeploymentItem(docFile, testFilesFolder)]
        public void HeavyDOCtoDOCX()
        {
            HeavyConversion<WordConverter>(docFile, FileTypes.doc, FileTypes.docx);
        }

        [TestMethod]
        [DeploymentItem(xlsFile, testFilesFolder)]
        public void HeavyXLStoXLSX()
        {
            HeavyConversion<ExcelConverter>(xlsFile, FileTypes.xls, FileTypes.xlsx);
        }

        [TestMethod]
        [DeploymentItem(pptFile, testFilesFolder)]
        public void HeavyPPTtoPPTX()
        {
            HeavyConversion<PowerPointConverter>(pptFile, FileTypes.ppt, FileTypes.pptx);
        }

        private void HeavyConversion<T>(string sourceFile, FileTypes sourceFormat, FileTypes targetFormat) where T: IConverter, new()
        {
            PooledConverter<T> pooledConverter = new PooledConverter<T>(convertersPoolSize);

            try
            {
                TestHighConcurrencyConversion(pooledConverter, sourceFile, sourceFormat, targetFormat, concurrentConversions);
            }
            finally
            {
                pooledConverter.Dispose();
            }
        }

    }
}
