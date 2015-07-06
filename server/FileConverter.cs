using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;

namespace LegacyOfficeConverter 
{
    public class FileConverter: IDisposable
    {
        private Microsoft.Office.Interop.Word.Application word;
        private Microsoft.Office.Interop.Excel.Application excel;
        private Microsoft.Office.Interop.PowerPoint.Application powerPoint;
        private OCRConsole OCRConvertor;


        private bool disposed = false;
        
        public FileConverter(string OCRConsolePath = null)
        {
            // Start Word
            word = new Microsoft.Office.Interop.Word.Application();
            word.Visible = false;

            // Start Excel
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = false;

            // Start Powerpoint
            powerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
            // Setting the Visible property like Word and Excel causes an exception.
            // The PowerPoint visibility is controlled using a parameter in the
            // document's open method.

            // Store the path of the OCRConsole
            OCRConvertor = new OCRConsole(OCRConsolePath);
        }

        /**
         * Converts the file at the provided path in a OOXML file, and returns the
         * resulting converted file path. The converted file is saved in the same
         * directory of the input file and with the same name, the only difference
         * is the extension.
         */
        public string Convert(string path)
        {
            string extension = Path.GetExtension(path).ToLower();
            string convertedPath = null;

            switch (extension)
            {
                case ".doc":
                case ".dot":
                case ".rtf":
                    convertedPath = convertWord(path);
                    break;

                case ".xls":
                case ".xlt":
                    convertedPath = convertExcel(path);
                    break;

                case ".ppt":
                case ".pps":
                case ".pot":
                    convertedPath = convertPowerpoint(path);
                    break;

                case ".pdf":
                case ".jpg":
                case ".tiff":
                case ".png":
                    convertedPath = convertOCR(path);
                    break;

                default:
                    // Unsupported exception
                    throw new Exception("FileConverter received an unsupported exception: " + extension + ".");
            }

            return convertedPath;
        }


        /**
         * Convert word file, return its converted path
         */ 
        private string convertWord(string path)
        {
            string convertedPath = null;
            int lastDotIndex = path.LastIndexOf('.');
            string pathWithoutExtension = path.Substring(0, lastDotIndex);
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
                    doc.SaveAs(FileName: pathWithoutExtension, FileFormat: WdSaveFormat.wdFormatDocumentDefault);
                    convertedPath = doc.FullName;
                }
                finally
                {
                    // Whatever happens, always release the COM object created for the document.
                    // .NET should handle COM objects release by itself, but I release them
                    // manually just to be sure. See http://goo.gl/7zv9Hj
                    if (doc != null)
                    {
                        doc.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(doc);
                        doc = null;
                    }
                }
            }
            return convertedPath;
        }


        /**
         * Convert Excel file, return its converted path
         */
        private string convertExcel(string path)
        {
            string convertedPath = null;
            int lastDotIndex = path.LastIndexOf('.');
            string pathWithoutExtension = path.Substring(0, lastDotIndex);
            lock (excel)
            {
                Workbook xls = null;
                try
                {
                    xls = excel.Workbooks.Open(Filename: path, ReadOnly: true);
                    if (xls == null)
                    {
                        throw new Exception("FileConverter could not open the file.");
                    }
                    xls.SaveAs(Filename: pathWithoutExtension, FileFormat: XlFileFormat.xlOpenXMLWorkbook);
                    convertedPath = xls.FullName;
                }
                finally
                {
                    if (xls != null)
                    {
                        xls.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(xls);
                        xls = null;
                    }
                }
            }
            return convertedPath;
        }


        /**
         * Convert powerpoint file, return its converted path
         */
        private string convertPowerpoint(string path)
        {
            string convertedPath = null;
            int lastDotIndex = path.LastIndexOf('.');
            string pathWithoutExtension = path.Substring(0, lastDotIndex);
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
                        ppt.Close();
                        Marshal.ReleaseComObject(ppt);
                        ppt = null;
                    }
                }
            }
            return convertedPath;
        }


        /**
         * Convert OCR file, return its converted path
         */
        private string convertOCR(string path)
        {
            return OCRConvertor.OCRrecognition(path);
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

            if (disposing) {
                // No managed resources to dispose
            }

            try
            {
                word.Quit(SaveChanges: false);
            } catch(Exception e) {
                Console.WriteLine("WARNING: exception while quitting Word instance. Full error:");
                Console.WriteLine(e);
            }
            Marshal.ReleaseComObject(word);

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

            disposed = true;
        }

        ~FileConverter()
        {
            Dispose(false);
        }
    }
}
