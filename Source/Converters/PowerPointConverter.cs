using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using Translated.MateCAT.WinConverter.ConversionServer;
using log4net;
using static System.Reflection.MethodBase;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class PowerPointConverter : IConverter, IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        private static int[] supportedFormats = { (int)FileTypes.ppt, (int)FileTypes.pps, (int)FileTypes.pot, (int)FileTypes.pptx, (int)FileTypes.pptm, (int)FileTypes.ppsx, (int)FileTypes.ppsm, (int)FileTypes.potx, (int)FileTypes.potm };

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
            log.Info("PowerPoint instance started");
            // Setting the Visible property like Word and Excel causes an exception.
            // The PowerPoint visibility is controlled using a parameter in the
            // document's open method.
            powerPoint.DisplayAlerts = PpAlertLevel.ppAlertsNone;
            powerPoint.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        }

        private void DestroyPowerPointInstance()
        {
            try
            {
                powerPoint.Quit();
                log.Info("PowerPoint instance destroyed");
            }
            catch (Exception e)
            {
                log.Warn("Exception while closing PowerPoint instance: check that the instance is properly closed", e);
            }

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
                catch (Exception e)
                {
                    // Obviously, in case of any error the instance is not working.
                    log.Warn("The PowerPoint instance is not working", e);
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
                        catch (Exception e)
                        {
                            log.Warn("Exception while closing test document", e);
                        }
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
                    log.Info("Word instance not working properly: restarting");
                    DestroyPowerPointInstance();
                    CreatePowerPointInstance();
                }

                Presentation ppt = null;
                try
                {
                    // Open the file
                    try
                    {
                        ppt = powerPoint.Presentations.Open2007(
                            FileName: sourceFilePath, 
                            ReadOnly: Microsoft.Office.Core.MsoTriState.msoTrue, 
                            WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse,
                            OpenAndRepair: Microsoft.Office.Core.MsoTriState.msoTrue);
                    }
                    catch (Exception e)
                    {
                        throw new BrokenSourceException("Exception opening source file.", e);
                    }
                    if (ppt == null)
                    {
                        throw new BrokenSourceException("Source file opened is null.");
                    }

                    // Select the target format
                    PpSaveAsFileType msOfficeTargetFormat;
                    switch (targetFormat)
                    {
                        case (int)FileTypes.ppt:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsPresentation;
                            break;
                        case (int)FileTypes.pps:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsShow;
                            break;
                        case (int)FileTypes.pot:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsTemplate;
                            break;
                        case (int)FileTypes.pptx:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsOpenXMLPresentation;
                            break;
                        case (int)FileTypes.pptm:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsOpenXMLPresentationMacroEnabled;
                            break;
                        case (int)FileTypes.ppsx:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsOpenXMLShow;
                            break;
                        case (int)FileTypes.ppsm:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsOpenXMLShowMacroEnabled;
                            break;
                        case (int)FileTypes.potx:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsOpenXMLTemplate;
                            break;
                        case (int)FileTypes.potm:
                            msOfficeTargetFormat = PpSaveAsFileType.ppSaveAsOpenXMLTemplateMacroEnabled;
                            break;
                        default:
                            throw new Exception("Unexpected target format");
                    }

                    try
                    {
                        // Save the file in the target format
                        ppt.SaveAs(FileName: targetFilePath, FileFormat: msOfficeTargetFormat);
                    }
                    catch (Exception e)
                    {
                        throw new ConversionException("Conversion exception.", e);
                    }

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
                        catch (Exception e)
                        {
                            log.Warn("Exception while closing the source document", e);
                        }
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
