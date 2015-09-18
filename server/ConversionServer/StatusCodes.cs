namespace Translated.MateCAT.LegacyOfficeConverter.ConversionServer
{
    public enum StatusCodes
    {
        Ok = 0,
        BadFileType = 1,
        BadFileSize = 2,
        BrokenFile = 3,
        ConvertedFileTooBig = 4,
        InternalServerError = 5,
        UnsupportedConversion = 7,
    }
}
