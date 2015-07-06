using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    class OCRConsole
    {

        private static readonly string[] validExtensions = { ".jpg", ".png", ".tiff", ".pdf" };
        private string toolPath;

        /// <summary>
        /// Console constructor, receiving the path to the OCR console
        /// </summary>
        /// <param name="toolPath"></param>
        public OCRConsole(string toolPath)
        {
            this.toolPath = toolPath;
        }
        
        
        /// <summary>
        /// Perform an OCR processing over the given file and output it in the same path
        /// </summary>
        /// <param name="filePath">File path</param>
        public string OCRrecognition(string filePath)
        {

            // Check that its a valid extension
            string extension = Path.GetExtension(filePath);
            if (!validExtensions.Contains(extension.ToLower()))
                throw new ArgumentException("FileConverter received an unsupported exception: " + extension + ".");

            // Obtain the out path and execute the console
            string outPath = Path.ChangeExtension(filePath, ".docx");
            RunConsole(filePath, outPath);

            // Return out path
            return outPath;

        }


        /// <summary>
        /// Execute the OCR console program
        /// </summary>
        /// <param name="input">Input file path</param>
        /// <param name="output">Output file path</param>
        private void RunConsole(string input, string output)
        {

            // Check that the console it is installed
            if (!File.Exists(toolPath))
                throw new Exception("The OCR is not installed in the server");

            // Set the process ready
            var process = new Process();
            process.StartInfo.Arguments = "\"" + input + "\" \"" + output + "\"";
            process.StartInfo.FileName = toolPath;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.RedirectStandardOutput = true;

            // Execute it
            var stdOutput = new StringBuilder();
            process.OutputDataReceived += (sender, args) => stdOutput.Append(args.Data);
            string stdError = null;
            process.Start();
            process.BeginOutputReadLine();
            stdError = process.StandardError.ReadToEnd();
            process.WaitForExit();


            // Receive errors (this should never happen)
            if (process.ExitCode != 0)
            {
                throw new Exception("It was not possible to OCR convert the file");

                /* ERROR HANDLING TODO
                var message = new StringBuilder();
                if (!string.IsNullOrEmpty(stdError))
                {
                    message.AppendLine(stdError);
                }

                if (stdOutput.Length != 0)
                {
                    message.AppendLine("Std output:");
                    message.AppendLine(stdOutput.ToString());
                }

                throw new Exception(" Process finished with exit code = " + process.ExitCode + ": " + message);
                */
            }
        }

    }
}
