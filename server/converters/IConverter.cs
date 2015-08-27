using System;

namespace LegacyOfficeConverter
{
    interface IConverter
    {
        /// <summary>
        /// Converts the file at the provided path in a OOXML file, and returns the
        /// resulting converted file path.The converted file is saved in the same
        /// directory of the input file and with the same name, the only difference
        /// is the extension.
        /// </summary
        void Convert(string inputPath, string outputPath);
    }
}
