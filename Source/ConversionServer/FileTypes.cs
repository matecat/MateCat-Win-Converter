namespace Translated.MateCAT.WinConverter.ConversionServer
{
    public enum FileTypes
    {

        // Word formats

        doc = 1,
        dot = 2, // Legacy Word Template

        docx = 3,
        docm = 4,

        dotx = 5,
        dotm = 6,

        rtf = 7,
       
        // Excel formats

        xls = 8,
        xlt = 9, // Legacy Excel Template

        xlsx = 10,
        xlsm = 11,

        xltm = 12,
        xltx = 13,

        // Powerpoint formats

        ppt = 14,
        pps = 15, // Legacy PowerPoint slideshow
        pot = 16, // Legacy PowerPoint Template

        pptx = 17,
        pptm = 18,

        ppsx = 19,
        ppsm = 20,

        potx = 21,
        potm = 22,

        // PDF & OCR

        pdf = 23,
        bmp = 24,
        gif = 25,
        png = 26,
        jpeg = 27,
        tiff = 28
    }
}
