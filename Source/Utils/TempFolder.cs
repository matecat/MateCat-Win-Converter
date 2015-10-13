using System;
using System.Configuration;
using System.IO;
using System.Linq;

namespace Translated.MateCAT.WinConverter.Utils
{
    public class TempFolder
    {
        private static readonly string rootTmpFolder;
        private static readonly object fileSystemLock = new object();

        static TempFolder() {
            rootTmpFolder = ConfigurationManager.AppSettings.Get("CachePath");
            
            // If specified cache path is empty, use the system default temp dir
            if (rootTmpFolder == null || rootTmpFolder == "")
            {
                rootTmpFolder = Path.GetTempPath();
            }

            // Path canonicalization
            rootTmpFolder = GetCanonicalPath(rootTmpFolder);
        }


        private readonly string dir;

        public TempFolder(int conversionId)
        {
            // This class will be heavily used in a threaded environment,
            // be careful and make it thread safe
            lock (fileSystemLock)
            {
                dir = rootTmpFolder + conversionId + Path.DirectorySeparatorChar;

                int duplicates = 0;
                while (Directory.Exists(dir))
                {
                    duplicates++;
                    dir = rootTmpFolder + conversionId + " (" + duplicates + ")" + Path.DirectorySeparatorChar;
                }
                Directory.CreateDirectory(dir);
            }
        }

        public TempFolder() : this((new Random()).Next()) { }

        public string getFilePath(string filename)
        {
            return dir + filename;
        }

        public override string ToString()
        {
            return dir;
        }


        // Thanks to http://stackoverflow.com/a/20406065
        private static string GetCanonicalPath(string path)
        {
            path = path
                .Trim()
                .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

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
