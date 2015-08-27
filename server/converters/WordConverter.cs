using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;
using System.IO;

namespace LegacyOfficeConverter
{
    class WordConverter : IConverter, IConverterSanityCheck, IDisposable
    {
        private Application word;

        private bool disposed = false;

        public WordConverter()
        {
            // Start Word
            word = new Application();
            word.Visible = false;
        }


        public bool isWorking()
        {
            lock (word)
            {
                Document doc = null;
                try
                {
                    doc = word.Documents.Add();
                    // Return true if the instance successfully created an empty document.
                    // The document will be gracefully closed and released in the finally block.
                    return (doc != null);
                }
                catch
                {
                    // Obviously, in case of any error the instance is not working.
                    return false;
                }
                finally
                {
                    // Whatever happens, always release the COM object created for the document.
                    // .NET should handle COM objects release by itself, but I release them
                    // manually just to be sure. See http://goo.gl/7zv9Hj
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                        }
                        catch { } // Skip any kind of error
                        finally
                        {
                            Marshal.ReleaseComObject(doc);
                            doc = null;
                        }
                    }
                }
            }
        }

        public void Convert(string inputPath, string outputPath)
        {
            lock (word)
            {
                Document doc = null;
                try
                {
                    doc = word.Documents.Open(FileName: inputPath, ReadOnly: true);
                    if (doc == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }
                    String outExtension = Path.GetExtension(outputPath).Replace(".","");
                    WdSaveFormat outFormat;
                    if (outExtension == "rtf")
                        outFormat = WdSaveFormat.wdFormatRTF;
                    //else if (outExtension == "doc")
                    //    outFormat = WdSaveFormat.wdFormatDocument97;
                    //else if (outExtension == "pdf")
                    //    outFormat = WdSaveFormat.wdFormatPDF;
                    else
                        outFormat = WdSaveFormat.wdFormatXMLDocument;
                    doc.SaveAs(FileName: outputPath, FileFormat: outFormat);
                }
                finally
                {
                    // Whatever happens, always release the COM object created for the document.
                    // .NET should handle COM objects release by itself, but I release them
                    // manually just to be sure. See http://goo.gl/7zv9Hj
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                        }
                        catch { }
                        finally
                        {
                            Marshal.ReleaseComObject(doc);
                            doc = null;
                        }
                    }
                }
            }
        }

        /*
         * Pay great attention to the dispose/destruction functions, it's
         * very important to release the used Office objects properly.
         */
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // No managed resources to dispose
            }

            try
            {
                word.Quit(SaveChanges: false);
            }
            catch { } // Ignore every exception

            Marshal.ReleaseComObject(word);
            word = null;

            disposed = true;
        }

        ~WordConverter()
        {
            Dispose(false);
        }
    }
}
