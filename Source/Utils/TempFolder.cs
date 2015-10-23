using log4net;
using System;
using System.Configuration;
using System.IO;
using static System.Reflection.MethodBase;

namespace Translated.MateCAT.WinConverter.Utils
{
    public class TempFolder
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        private static readonly string rootTmpFolder;
        private static readonly string rootErrorsFolder;
        private static readonly object fileSystemLock = new object();

        private static readonly Random random = new Random();

        static TempFolder() {
            rootTmpFolder = ConfigurationManager.AppSettings.Get("CachePath");

            // If specified cache path is empty, use the system default temp dir
            if (rootTmpFolder == null || rootTmpFolder.Trim() == "")
            {
                rootTmpFolder = Path.GetTempPath();
            }

            // Path canonicalization
            rootTmpFolder = GetCanonicalPath(rootTmpFolder);


            rootErrorsFolder = ConfigurationManager.AppSettings.Get("ErrorsPath");

            if (rootErrorsFolder.Trim() == "")
            {
                rootErrorsFolder = null;
            }
            else
            {
                rootErrorsFolder = GetCanonicalPath(rootErrorsFolder);
            }
        }

        private static int GetRandomNumberThreadSafe()
        {
            lock(random)
            {
                return random.Next();
            }
        }


        private readonly string dir;
        private readonly int conversionId;

        public TempFolder(int conversionId)
        {
            // This class will be heavily used in a threaded environment,
            // be careful and make it thread safe
            lock (fileSystemLock)
            {
                this.conversionId = conversionId;
                dir = rootTmpFolder + conversionId + Path.DirectorySeparatorChar;
                Directory.CreateDirectory(dir);
            }
        }

        public TempFolder() : this(GetRandomNumberThreadSafe()) { }

        public string getFilePath(string filename)
        {
            return dir + filename;
        }

        public void Release(bool error = false)
        {
            if (error && rootErrorsFolder != null)
            {
                string errorDir = rootErrorsFolder 
                    + DateTime.Now.ToString("yyyy-MM-dd") + Path.DirectorySeparatorChar
                    + DateTime.Now.ToString("HH-mm") + "-" + conversionId;
                try
                {
                    MoveDirectory(dir, errorDir);
                    log.Info("folder with temp files moved to " + errorDir);
                }
                catch (Exception e)
                {
                    log.Error("exception while moving temp folder to errors folder; will try to copy it", e);
                    try
                    {
                        CopyDirectory(dir, errorDir);
                        log.Info("folder with temp files copied to " + errorDir);
                    }
                    catch (Exception ee)
                    {
                        log.Error("exception while copying temp folder to errors folder", ee);
                    }
                }
            }

            // Directory can exist also because the MoveDirectory in the previous block failed.
            if (Directory.Exists(dir))
            {
                try
                {
                    Directory.Delete(dir, true);
                }
                catch (Exception e)
                {
                    log.Error("exception while deleting temp folder", e);
                }
            }
        }

        public override string ToString()
        {
            return dir;
        }


        /// <summary>
        /// Path sanitization: trim, replace of slashes, ensures ALWAYS ending with a backslash.
        /// Thanks to http://stackoverflow.com/a/20406065
        /// </summary>
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

        /// <summary>
        /// Regular .NET methods don't allow moving folders in different volumes.
        /// This function does.
        /// Thanks to http://stackoverflow.com/a/32997504
        /// </summary>
        private static void MoveDirectory(string source, string target)
        {
            DirectoryInfo sourceInfo = new DirectoryInfo(source);
            DirectoryInfo targetInfo = new DirectoryInfo(target);

            if (!targetInfo.Exists)
                targetInfo.Create();

            foreach (var file in sourceInfo.GetFiles())
                file.MoveTo(Path.Combine(targetInfo.FullName, file.Name));

            foreach (var subdir in sourceInfo.GetDirectories())
                MoveDirectory(subdir.FullName, targetInfo.CreateSubdirectory(subdir.Name).FullName);
        }

        /// <summary>
        /// Regular .NET methods don't allow copying folders in different volumes.
        /// This function does.
        /// Thanks to http://stackoverflow.com/a/32997504
        /// </summary>
        private static void CopyDirectory(string source, string target)
        {
            DirectoryInfo sourceInfo = new DirectoryInfo(source);
            DirectoryInfo targetInfo = new DirectoryInfo(target);

            if (!targetInfo.Exists)
                targetInfo.Create();

            foreach (var file in sourceInfo.GetFiles())
                file.CopyTo(Path.Combine(targetInfo.FullName, file.Name));

            foreach (var subdir in sourceInfo.GetDirectories())
                CopyDirectory(subdir.FullName, targetInfo.CreateSubdirectory(subdir.Name).FullName);
        }
    }
}
