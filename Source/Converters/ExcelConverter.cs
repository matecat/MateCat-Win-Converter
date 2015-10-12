using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using Translated.MateCAT.WinConverter.ConversionServer;
using log4net;
using static System.Reflection.MethodBase;

namespace Translated.MateCAT.WinConverter.Converters
{
    public class ExcelConverter : IConverter, IDisposable
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        private static int[] supportedFormats = { (int)FileTypes.xls, (int)FileTypes.xlt, (int)FileTypes.xlsx, (int)FileTypes.xlsm, (int)FileTypes.xltx, (int)FileTypes.xltm };

        private readonly object lockObj = new object();
        private Application excel;

        private bool disposed = false;

        public ExcelConverter()
        {
            CreateExcelInstance();
        }

        private void CreateExcelInstance()
        {
            excel = new Application();
            log.Info("Excel instance started");
            excel.Visible = false;
            excel.DisplayAlerts = false;
        }

        private void DestroyExcelInstance()
        {
            try
            {
                excel.Quit();
                log.Info("Excel instance destroyed");
            }
            catch (Exception e)
            {
                log.Warn("Exception while closing Excel instance: check that the instance is properly closed", e);
            }

            Marshal.ReleaseComObject(excel);
            excel = null;
        }

        private bool IsExcelWorking()
        {
            lock (lockObj)
            {
                Workbook xls = null;
                try
                {
                    xls = excel.Workbooks.Add();
                    // Return true if the instance successfully created an empty document.
                    // The document will be gracefully closed and released in the finally block.
                    return (xls != null);
                }
                catch (Exception e)
                {
                    // Obviously, in case of any error the instance is not working.
                    log.Warn("The Excel instance is not working", e);
                    return false;
                }
                finally
                {
                    // Whatever happens, always release the COM object created for the document.
                    // .NET should handle COM objects release by itself, but I release them
                    // manually just to be sure. See http://goo.gl/7zv9Hj
                    if (xls != null)
                    {
                        try
                        {
                            xls.Close(SaveChanges: false);
                        }
                        catch (Exception e)
                        {
                            log.Warn("Exception while closing test document", e);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(xls);
                            xls = null;
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

            // Covnersion supported, do it
            lock (lockObj)
            {
                // Ensure Excel instance is working
                if (!IsExcelWorking())
                {
                    log.Info("Excel instance not working properly: restarting");
                    DestroyExcelInstance();
                    CreateExcelInstance();
                }

                Workbook xls = null;
                try
                {
                    // Open the file
                    try
                    {
                        xls = excel.Workbooks.Open(Filename: sourceFilePath, ReadOnly: true);
                    }
                    catch (Exception e)
                    {
                        throw new BrokenSourceException("Exception opening source file.", e);
                    }
                    if (xls == null)
                    {
                        throw new BrokenSourceException("Source file opened is null.");
                    }

                    // Select the target format
                    XlFileFormat msOfficeTargetFormat;
                    switch (targetFormat)
                    {
                        case (int)FileTypes.xls:
                            msOfficeTargetFormat = XlFileFormat.xlExcel8;
                            break;
                        case (int)FileTypes.xlt:
                            msOfficeTargetFormat = XlFileFormat.xlTemplate8;
                            break;
                        case (int)FileTypes.xlsx:
                            msOfficeTargetFormat = XlFileFormat.xlOpenXMLWorkbook;
                            break;
                        case (int)FileTypes.xlsm:
                            msOfficeTargetFormat = XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                            break;
                        case (int)FileTypes.xltx:
                            msOfficeTargetFormat = XlFileFormat.xlOpenXMLTemplate;
                            break;
                        case (int)FileTypes.xltm:
                            msOfficeTargetFormat = XlFileFormat.xlOpenXMLTemplateMacroEnabled;
                            break;
                        default:
                            throw new Exception("Unexpected target format");
                    }

                    try
                    {
                        // Save the file in the target format
                        xls.SaveAs(
                            Filename: targetFilePath, 
                            FileFormat: msOfficeTargetFormat, 
                            ConflictResolution: XlSaveConflictResolution.xlLocalSessionChanges);
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
                    if (xls != null)
                    {
                        try
                        {
                            xls.Close(SaveChanges: false);
                        }
                        catch (Exception e)
                        {
                            log.Warn("Exception while closing the source document", e);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(xls);
                            xls = null;
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

            DestroyExcelInstance();

            disposed = true;
        }

        ~ExcelConverter()
        {
            Dispose(false);
        }
    }
}
