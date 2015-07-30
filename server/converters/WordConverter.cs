using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;
using System.IO;

namespace LegacyOfficeConverter
{
    class WordConverter : IConverter, IDisposable
    {
        private Application word;

        private bool disposed = false;

        public WordConverter()
        {
            // Start Word
            word = new Application();
            word.Visible = false;
        }

        public string Convert(string path)
        {
            int lastDotIndex = path.LastIndexOf('.');
            string pathWithoutExtension = path.Substring(0, lastDotIndex);
            
            string convertedPath = null;
            lock (word)
            {
                Document doc = null;
                try
                {
                    doc = word.Documents.Open(FileName: path, ReadOnly: true);
                    if (doc == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }
                    //doc.SaveAs(FileName: pathWithoutExtension, FileFormat: WdSaveFormat.wdFormatDocumentDefault);
                    doc.SaveAs(FileName: pathWithoutExtension, FileFormat: WdSaveFormat.wdFormatXMLDocument);
                    convertedPath = doc.FullName;
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
            return convertedPath;
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
            catch (Exception e)
            {
                Console.WriteLine("WARNING: exception while quitting Word instance. Full error:");
                Console.WriteLine(e);
            }
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
