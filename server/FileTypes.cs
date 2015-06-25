using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    enum FileTypes
    {
        doc = 1,
        xls = 2,
        ppt = 3,
        dot = 4, // Legacy Word Template
        xlt = 5, // Legacy Excel Template
        pot = 6, // Legacy PowerPoint Template
        pps = 7, // Legacy PowerPoint slideshow
        rtf = 8
    }
}
