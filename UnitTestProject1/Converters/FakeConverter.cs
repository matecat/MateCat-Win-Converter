using System.IO;

namespace Translated.MateCAT.WinConverter.Converters
{
    class FakeConverter : IConverter
    {
        public bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat)
        {
            File.WriteAllBytes(targetFilePath, new byte[] { 0 });
            return true;
        }
    }
}
