using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    public class ExcelConverter : IConverter, IConverterSanityCheck, IDisposable
    {
        private static int[] supportedFormats = { (int)FileTypes.xls, (int)FileTypes.xlsx };

        private Application excel;

        private bool disposed = false;

        public ExcelConverter()
        {
            // Start Excel
            excel = new Application();
            excel.Visible = false;
        }

        public bool isWorking()
        {
            lock (excel)
            {
                Workbook xls = null;
                try
                {
                    xls = excel.Workbooks.Add();
                    // Return true if the instance successfully created an empty document.
                    // The document will be gracefully closed and released in the finally block.
                    return (xls != null);
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
                    if (xls != null)
                    {
                        try
                        {
                            xls.Close(SaveChanges: false);
                        }
                        catch { } // Skip any kind of error
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
            lock (excel)
            {
                Workbook xls = null;
                try
                {
                    // Open the file
                    xls = excel.Workbooks.Open(Filename: sourceFilePath, ReadOnly: true);
                    if (xls == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }

                    // Select the target format
                    XlFileFormat msOfficeTargetFormat;
                    switch (targetFormat)
                    {
                        case (int)FileTypes.xlsx:
                            msOfficeTargetFormat = XlFileFormat.xlOpenXMLWorkbook;
                            break;
                        case (int)FileTypes.xls:
                            msOfficeTargetFormat = XlFileFormat.xlExcel8;
                            break;
                        default:
                            throw new Exception("Unexpected target format");
                    }

                    // Save the file in the target format
                    xls.SaveAs(Filename: targetFilePath, FileFormat: msOfficeTargetFormat);

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
                        catch { }
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

            try
            {
                excel.Quit();
            }
            catch { } // Ignore every exception

            Marshal.ReleaseComObject(excel);
            excel = null;

            disposed = true;
        }

        ~ExcelConverter()
        {
            Dispose(false);
        }
    }
}
