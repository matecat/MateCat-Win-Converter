using Microsoft.Office.Interop.Word;
using System;
using System.Runtime.InteropServices;
using System.Linq;
using Translated.MateCAT.WinConverter.ConversionServer;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class WordConverter : IConverter, IDisposable
    {
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
            word.Visible = false;
            word.DisplayAlerts = WdAlertLevel.wdAlertsNone;
        }

        private void DestroyWordInstance()
        {
            try
            {
                word.Quit(SaveChanges: false);
            }
            catch { } // Ignore every exception

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
                    DestroyWordInstance();
                    CreateWordInstance();
                }

                Document doc = null;
                try
                {
                    // Open the file
                    doc = word.Documents.Open(FileName: sourceFilePath, ReadOnly: true);
                    if (doc == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
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

                    // Save the file in the target format
                    doc.SaveAs(FileName: targetFilePath, FileFormat: msOfficeTargetFormat);
                    
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

            DestroyWordInstance();

            disposed = true;
        }

        ~WordConverter()
        {
            Dispose(false);
        }
    }
}
