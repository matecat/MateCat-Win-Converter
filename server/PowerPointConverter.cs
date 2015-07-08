using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Runtime.InteropServices;

namespace LegacyOfficeConverter
{
    class PowerPointConverter : IConverter
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

        public string Convert(string path)
        {
            int lastDotIndex = path.LastIndexOf('.');
            string pathWithoutExtension = path.Substring(0, lastDotIndex);

            string convertedPath = null;
            lock (powerPoint)
            {
                Presentation ppt = null;
                try
                {
                    ppt = powerPoint.Presentations.Open(FileName: path, ReadOnly: MsoTriState.msoTrue, WithWindow: MsoTriState.msoFalse);
                    if (ppt == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }
                    ppt.SaveAs(FileName: pathWithoutExtension, FileFormat: PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                    convertedPath = ppt.FullName;
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
                powerPoint.Quit();
            }
            catch (Exception e)
            {
                Console.WriteLine("WARNING: exception while quitting PowerPoint instance. Full error:");
                Console.WriteLine(e);
            }
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
