using System.Configuration;
using System.IO;

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

        public TempFolder()
        {
            // This class will be heavily used in a threaded environment,
            // be careful and make it thread safe
            lock (fileSystemLock)
            {
                do
                {
                    dir = rootTmpFolder + Path.GetRandomFileName() + Path.DirectorySeparatorChar;
                }
                while (Directory.Exists(dir));
                Directory.CreateDirectory(dir);
            }
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
