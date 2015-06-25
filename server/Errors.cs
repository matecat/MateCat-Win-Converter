using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    public enum Errors
    {
        BadFileType = 1,
        BadFileSize = 2,
        BrokenFile = 3,
        ConvertedFileTooBig = 4,
        InternalServerError = 5
    }
}
