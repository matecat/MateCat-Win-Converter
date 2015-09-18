using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using static Translated.MateCAT.LegacyOfficeConverter.Converters.ConversionTestHelper;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    [TestClass]
    public class PooledConverterTest
    {
        const int parallelism = 10;
        const int conversionQueueSize = 30;

        static PooledConverter<WordConverter> converter;

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            converter = new PooledConverter<WordConverter>(parallelism);
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            converter.Dispose();
        }

        [TestMethod]
        [DeploymentItem(docFile, testFilesFolder)]
        public void HeavyLoadDOCtoDOCX()
        {
            string sourceFile = Path.GetFullPath(docFile);

            int successes = 0;
            List<Thread> threads = new List<Thread>();
            for (int i = 0; i < conversionQueueSize; i++)
            {
                Thread thread = new Thread(delegate () {
                    string targetFile = Path.GetTempFileName();
                    try
                    {
                        bool converted = converter.Convert(sourceFile, (int)FileTypes.doc, targetFile, (int)FileTypes.docx);
                        if (converted) Interlocked.Increment(ref successes);
                    }
                    catch (Exception e) {
                        Console.WriteLine(e);
                    }
                    finally
                    {
                        if (File.Exists(targetFile)) File.Delete(targetFile);                        
                    }
                });
                threads.Add(thread);
            }
            foreach (var thread in threads)
            {
                thread.Start();
            }
            foreach (var thread in threads)
            {
                Console.WriteLine("ended");
                thread.Join();
            }
            Console.WriteLine(successes);
            Assert.AreEqual(conversionQueueSize, successes);
        }

    }
}
