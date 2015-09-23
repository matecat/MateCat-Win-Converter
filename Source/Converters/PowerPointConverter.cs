using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using Translated.MateCAT.WinConverter.ConversionServer;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class PowerPointConverter : IConverter, IDisposable
    {
        private static int[] supportedFormats = { (int)FileTypes.pptx, (int)FileTypes.ppt };

        private readonly object lockObj = new object();
        private Application powerPoint;

        private bool disposed = false;

        public PowerPointConverter()
        {
            CreatePowerPointInstance();
        }

        private void CreatePowerPointInstance()
        {
            // Start Powerpoint
            powerPoint = new Application();
            // Setting the Visible property like Word and Excel causes an exception.
            // The PowerPoint visibility is controlled using a parameter in the
            // document's open method.
            powerPoint.DisplayAlerts = PpAlertLevel.ppAlertsNone;
        }

        private void DestroyPowerPointInstance()
        {
            try
            {
                powerPoint.Quit();
            }
            catch { } // Ignore every exception

            Marshal.ReleaseComObject(powerPoint);
            powerPoint = null;
        }


        private bool IsPowerPointWorking()
        {
            lock (lockObj)
            {
                Presentation ppt = null;
                try
                {
                    ppt = powerPoint.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);
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

        public bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat)
        {
            // Check if the required conversion is supported
            if (!supportedFormats.Contains(sourceFormat) || !supportedFormats.Contains(targetFormat))
            {
                return false;
            }

            // Conversion supported, do it
            lock (lockObj)
            {
                // Ensure PowerPoint instance is working
                if (!IsPowerPointWorking())
                {
                    DestroyPowerPointInstance();
                    CreatePowerPointInstance();
                }

                Presentation ppt = null;
                try
                {
                    // Open the file
                    ppt = powerPoint.Presentations.Open(FileName: sourceFilePath, ReadOnly: Microsoft.Office.Core.MsoTriState.msoTrue, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
                    if (ppt == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }

                    // Select the target format
                    PpSaveAsFileType msOfficeTargetFormat;
                    switch (targetFormat)
                    {
                        case (int)FileTypes.pptx:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsOpenXMLPresentation;
                            break;
                        case (int)FileTypes.ppt:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsPresentation;
                            break;
                        default:
                            throw new Exception("Unexpected target format");
                    }

                    // Save the file in the target format
                    ppt.SaveAs(FileName: targetFilePath, FileFormat: msOfficeTargetFormat);

                    // Everything ok, return the success to the caller
                    return true;
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

            DestroyPowerPointInstance();

            disposed = true;
        }

        ~PowerPointConverter()
        {
            Dispose(false);
        }
    }
}
