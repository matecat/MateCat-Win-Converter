namespace Translated.MateCAT.WinConverter.ConversionServer
{
    public enum StatusCodes
    {
        Ok,  // => 0
        BadFileType,
        BadFileSize,
        BrokenSourceFile,
        ConvertedFileTooBig,
        ConversionError,
        InternalServerError,
        UnsupportedConversion // => 7
    }
}
