using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Translated.MateCAT.LegacyOfficeConverter.ConversionServer
{
    public enum FileTypes
    {
        docx = 0,
        doc = 1,
        xls = 2,
        ppt = 3,
        dot = 4, // Legacy Word Template
        xlt = 5, // Legacy Excel Template
        pot = 6, // Legacy PowerPoint Template
        pps = 7, // Legacy PowerPoint slideshow
        rtf = 8,
        pdf = 9,
        png = 10,
        jpg = 11,
        tiff = 12,
        xlsx = 13,
        pptx = 14
    }
}
