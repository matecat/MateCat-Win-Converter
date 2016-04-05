using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;
using System.Linq;
using Translated.MateCAT.WinConverter.ConversionServer;
using log4net;
using static System.Reflection.MethodBase;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class WordConverter : IConverter, IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        private static int[] supportedFormats = { (int)FileTypes.doc, (int)FileTypes.dot, (int)FileTypes.docx, (int)FileTypes.docm, (int)FileTypes.dotx, (int)FileTypes.dotm, (int)FileTypes.rtf };

        private readonly object lockObj = new object();
        private Application word;

        private bool disposed = false;

        public WordConverter()
        {
            CreateWordInstance();
        }

        private void CreateWordInstance()
        {
            word = new Application();
            log.Info("Word instance started");
            word.Visible = false;
            word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            word.ScreenUpdating = false;
            word.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
        }

        private void DestroyWordInstance()
        {
            try
            {
                word.Quit(SaveChanges: false);
                log.Info("Word instance destroyed");
            }
            catch (Exception e)
            {
                log.Warn("Exception while closing Word instance: check that the instance is properly closed", e);
            }

            Marshal.ReleaseComObject(word);
            word = null;
        }

        private bool IsWordWorking()
        {
            lock (lockObj)
            {
                Document doc = null;
                try
                {
                    doc = word.Documents.Add();
                    // Return true if the instance successfully created an empty document.
                    // The document will be gracefully closed and released in the finally block.
                    return (doc != null);
                }
                catch (Exception e)
                {
                    // Obviously, in case of any error the instance is not working.
                    log.Warn("The Word instance is not working", e);
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
                // Ensure Word instance is working
                if (!IsWordWorking())
                {
                    log.Info("Word instance not working properly: restarting");
                    DestroyWordInstance();
                    CreateWordInstance();
                }

                Document doc = null;
                try
                {
                    // Open the file
                    try
                    {
                        doc = word.Documents.Open(FileName: sourceFilePath, ConfirmConversions: false, ReadOnly: true, 
                            PasswordDocument: "-", PasswordTemplate: "-", WritePasswordDocument: "-", WritePasswordTemplate: "-",
                            Visible: false, OpenAndRepair: true, NoEncodingDialog: true);
                    }
                    catch (Exception e)
                    {
                        throw new BrokenSourceException("Exception opening source file.", e);
                    }
                    if (doc == null)
                    {
                        throw new BrokenSourceException("Source file opened is null.");
                    }

                    // Select the target format
                    WdSaveFormat msOfficeTargetFormat;
                    switch (targetFormat)
                    {
                        case (int)FileTypes.doc:
                            msOfficeTargetFormat = WdSaveFormat.wdFormatDocument97;
                            break;
                        case (int)FileTypes.dot:
                            msOfficeTargetFormat = WdSaveFormat.wdFormatTemplate97;
                            break;
                        case (int)FileTypes.docx:
                            msOfficeTargetFormat = WdSaveFormat.wdFormatXMLDocument;
                            break;
                        case (int)FileTypes.docm:
                            msOfficeTargetFormat = WdSaveFormat.wdFormatXMLDocumentMacroEnabled;
                            break;
                        case (int)FileTypes.dotx:
                            msOfficeTargetFormat = WdSaveFormat.wdFormatXMLTemplate;
                            break;
                        case (int)FileTypes.dotm:
                            msOfficeTargetFormat = WdSaveFormat.wdFormatXMLTemplateMacroEnabled;
                            break;
                        case (int)FileTypes.rtf:
                            msOfficeTargetFormat = WdSaveFormat.wdFormatRTF;
                            break;
                        default:
                            throw new Exception("Unexpected target format");
                    }

                    try
                    {
                        // Save the file in the target format
                        doc.SaveAs(FileName: targetFilePath, FileFormat: msOfficeTargetFormat, AddBiDiMarks: true);
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
                    // Whatever happens, always release the COM object created for the document.
                    // .NET should handle COM objects release by itself, but I release them
                    // manually just to be sure. See http://goo.gl/7zv9Hj
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                        }
                        catch (Exception e)
                        {
                            log.Warn("Exception while closing the source document", e);
                        }
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

            DestroyWordInstance();

            disposed = true;
        }

        ~WordConverter()
        {
            Dispose(false);
        }
    }
}
