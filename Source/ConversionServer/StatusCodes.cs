namespace Translated.MateCAT.WinConverter.ConversionServer
{
    public enum StatusCodes
    {
        Ok,  // => 0
        BadFileType,
        BadFileSize,
        BrokenSourceFile,
        ConvertedFileTooBig,
        InternalServerError,
        UnsupportedConversion // => 6
    }
}
