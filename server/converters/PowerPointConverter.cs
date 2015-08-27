using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Runtime.InteropServices;

namespace LegacyOfficeConverter
{
    class PowerPointConverter : IConverter, IConverterSanityCheck, IDisposable
    {
        private Application powerPoint;

        private bool disposed = false;

        public PowerPointConverter()
        {
            // Start Powerpoint
            powerPoint = new Application();
            // Setting the Visible property like Word and Excel causes an exception.
            // The PowerPoint visibility is controlled using a parameter in the
            // document's open method.
        }

        public bool isWorking()
        {
            lock (powerPoint)
            {
                Presentation ppt = null;
                try
                {
                    ppt = powerPoint.Presentations.Add(MsoTriState.msoFalse);
                    // Return true if the instance successfully created an empty document.
                    // The document will be gracefully closed and released in the finally block.
                    return (ppt != null);
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
                    if (ppt != null)
                    {
                        try
                        {
                            ppt.Close();
                        }
                        catch { } // Skip any kind of error
                        finally
                        {
                            Marshal.ReleaseComObject(ppt);
                            ppt = null;
                        }
                    }
                }
            }
        }

        public void Convert(string inputPath, string outputPath)
        {
            lock (powerPoint)
            {
                Presentation ppt = null;
                try
                {
                    ppt = powerPoint.Presentations.Open(FileName: inputPath, ReadOnly: MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);
                    if (ppt == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }
                    ppt.SaveAs(FileName: outputPath, FileFormat: PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                }
                finally
                {
                    if (ppt != null)
                    {
                        try
                        {
                            ppt.Close();
                        }
                        catch { }
                        finally
                        {
                            Marshal.ReleaseComObject(ppt);
                            ppt = null;
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
                powerPoint.Quit();
            }
            catch { } // Ignore every exception

            Marshal.ReleaseComObject(powerPoint);
            powerPoint = null;

            disposed = true;
        }

        ~PowerPointConverter()
        {
            Dispose(false);
        }
    }
}
