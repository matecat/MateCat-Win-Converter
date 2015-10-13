using log4net;
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using static System.Reflection.MethodBase;
using Translated.MateCAT.WinConverter.Converters;
using Translated.MateCAT.WinConverter.Utils;

namespace Translated.MateCAT.WinConverter.ConversionServer
{
    public class ConversionRequest
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        public const int BufferSize = 8192;

        private readonly Socket socket;
        private readonly IConverter fileConverter;

        public ConversionRequest(Socket socket, IConverter fileConverter)
        {
            this.socket = socket;
            this.fileConverter = fileConverter;
        }

        public void Run()
        {
            int bytesRead, bytesSent;
            byte[] buffer = new byte[BufferSize];

            log.Info("new request received");

            TempFolder tempFolder = null;
            string sourceFilePath = null;
            FileStream sourceFileStream = null;
            string targetFilePath = null;
            FileStream targetFileStream = null;

            try
            {
                // 1) Read the source file type

                int sourceFileTypeValue = ReceiveInt();
                FileTypes sourceFileType = (FileTypes)sourceFileTypeValue;
                log.Info("received source file type: " + sourceFileTypeValue + " ("+ sourceFileType +")");
                if (!Enum.IsDefined(typeof(FileTypes), sourceFileType))
                {
                    throw new ProtocolException(StatusCodes.BadFileType);
                }


                // 2) Read the target file type

                int targetFileTypeValue = ReceiveInt();
                FileTypes targetFileType = (FileTypes)targetFileTypeValue;
                log.Info("received target file type: " + targetFileTypeValue + " (" + targetFileType + ")");
                if (!Enum.IsDefined(typeof(FileTypes), targetFileType))
                {
                    throw new ProtocolException(StatusCodes.BadFileType);
                }


                // 3) Read the source file size

                int fileSize = ReceiveInt();
                log.Info("received source file size: " + fileSize);
                if (fileSize <= 0)
                {
                    throw new ProtocolException(StatusCodes.BadFileSize);
                }

                // 4) Read the source file and cache it on disk

                tempFolder = new TempFolder();
                sourceFilePath = tempFolder.getFilePath("source." + sourceFileType);
                sourceFileStream = new FileStream(sourceFilePath, FileMode.Create);
                int totalBytesRead = 0;
                log.Info("receiving source file...");
                while (totalBytesRead < fileSize)
                {
                    bytesRead = socket.Receive(buffer, BufferSize, 0);
                    sourceFileStream.Write(buffer, 0, bytesRead);
                    totalBytesRead += bytesRead;
                }
                sourceFileStream.Close();
                log.Info("source file received");


                // 5) Convert the source file
                bool converted = false;
                try
                {
                    targetFilePath = tempFolder.getFilePath("target." + targetFileType);
                    converted = fileConverter.Convert(sourceFilePath, (int)sourceFileType, targetFilePath, (int)targetFileType);
                }
                catch (BrokenSourceException e)
                {
                    throw new ProtocolException(StatusCodes.BrokenSourceFile, e);
                }
                catch (ConversionException e)
                {
                    throw new ProtocolException(StatusCodes.ConversionError, e);
                }
                if (!converted)
                {
                    throw new ProtocolException(StatusCodes.UnsupportedConversion);
                }


                // 6) Send back the ok status code

                SendInt((int)StatusCodes.Ok);
                log.Info("sent status code: 0 (ok!)");


                // 7) Send back the converted file size

                int filesize = 0;
                try
                {
                    filesize = (int)new FileInfo(targetFilePath).Length;
                }
                catch (OverflowException e)
                {
                    throw new ProtocolException(StatusCodes.ConvertedFileTooBig, e);
                }
                SendInt(filesize);
                log.Info("sent target file size: " + filesize);


                // 8) Send back the converted file

                buffer = new byte[BufferSize];
                targetFileStream = new FileStream(targetFilePath, FileMode.Open);
                log.Info("sending target file...");
                while (true)
                {
                    bytesRead = targetFileStream.Read(buffer, 0, BufferSize);
                    if (bytesRead == 0)
                    {
                        break;
                    }
                    else
                    {
                        bytesSent = socket.Send(buffer, bytesRead, 0);
                    }
                }
                log.Info("target file sent");
                targetFileStream.Close();

                // If execution arrives here, everything went well!
            }
            catch (ProtocolException e)
            {
                try
                {
                    log.Error("Protocol exception", e);
                    SendInt((int)e.statusCode);
                    log.Info("sent status code: " + (int)e.statusCode + " (" + e.statusCode + ")");
                }
                catch { }
            }
            catch (Exception e)
            {
                log.Error("General Exception", e);
                try
                {
                    SendInt((int)StatusCodes.InternalServerError);
                    log.Info("sent status code: " + (int)StatusCodes.InternalServerError + " (" + StatusCodes.InternalServerError + ")");
                }
                catch { }
            }
            finally
            {
                // Close streams                
                if (sourceFileStream != null) sourceFileStream.Close();
                if (targetFileStream != null) targetFileStream.Close();
                
                // Delete temp folder
                if (tempFolder != null)
                {
                    try
                    {
                        Directory.Delete(tempFolder.ToString(), true);
                    }
                    catch { }
                }

                EndPoint remoteEndPoint = socket.RemoteEndPoint;
                socket.Shutdown(SocketShutdown.Both);
                // Wait that client closes connection
                //socket.Receive(new byte[1], SocketFlags.Peek);
                socket.Close();
                log.Info("connection closed");
            }
        }

        private int ReceiveInt()
        {
            byte[] buffer = new byte[4];
            int bytesRead = socket.Receive(buffer, sizeof(int), 0);
            // The "network" value
            int netValue = BitConverter.ToInt32(buffer, 0);
            // The "windows" value
            int winValue = IPAddress.NetworkToHostOrder(netValue);
            return winValue;
        }

        private void SendInt(int value)
        {
            int netValue = IPAddress.HostToNetworkOrder(value);
            byte[] buffer = BitConverter.GetBytes(netValue);
            int bytesSent = socket.Send(buffer, sizeof(int), 0);
        }
    }
}
