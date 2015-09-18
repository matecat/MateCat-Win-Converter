using System.Configuration;
using System.IO;

namespace Translated.MateCAT.LegacyOfficeConverter.Utils
{
    class TempFolder
    {
        static readonly string rootTmpFolder;        

        static TempFolder() {
            rootTmpFolder = ConfigurationManager.AppSettings.Get("CachePath");
            // If specified cache path is empty, use the system default temp dir
            if (rootTmpFolder == null || rootTmpFolder == "")
            {
                rootTmpFolder = Path.GetTempPath();
            }
            rootTmpFolder = EnsureCorrectPath(rootTmpFolder);
        }


        private readonly string dir;

        public TempFolder()
        {
            dir = rootTmpFolder + Path.GetRandomFileName() + Path.DirectorySeparatorChar;
            Directory.CreateDirectory(dir);
        }

        public string getFilePath(string filename)
        {
            return dir + filename;
        }

        public override string ToString()
        {
            return dir;
        }


        // Thanks to http://stackoverflow.com/a/20406065
        private static string EnsureCorrectPath(string path)
        {
            path = path.Trim();
            path = path.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

            if (path.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                return path;
            }
            else
            {
                return path + Path.DirectorySeparatorChar;
            }
        }
    }
}
