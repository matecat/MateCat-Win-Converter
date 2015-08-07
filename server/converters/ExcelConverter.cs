using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;

namespace LegacyOfficeConverter
{
    class ExcelConverter : IConverter, IDisposable
    {
        private Application excel;

        private bool disposed = false;

        public ExcelConverter()
        {
            // Start Excel
            excel = new Application();
            excel.Visible = false;
        }

        public void Convert(string inputPath, string outputPath)
        {
            lock (excel)
            {
                Workbook xls = null;
                try
                {
                    xls = excel.Workbooks.Open(Filename: inputPath, ReadOnly: true);
                    if (xls == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }
                    xls.SaveAs(Filename: outputPath, FileFormat: XlFileFormat.xlOpenXMLWorkbook);
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
            catch (Exception e)
            {
                Console.WriteLine("WARNING: exception while quitting Excel instance. Full error:");
                Console.WriteLine(e);
            }
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
