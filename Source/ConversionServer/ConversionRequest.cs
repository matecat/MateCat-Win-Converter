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

        public const int BufferSize = 10*1024*1024; // 10 MB
        public const int SocketTimeout = 5000; // in milliseconds

        private readonly Socket socket;
        private readonly IConverter fileConverter;

        public ConversionRequest(Socket socket, IConverter fileConverter)
        {
            this.socket = socket;
            this.fileConverter = fileConverter;
        }

        public void Run()
        {
            socket.ReceiveTimeout = SocketTimeout;
            socket.SendTimeout = SocketTimeout;

            bool healthCheck = false;

            int bytesRead, bytesSent;
            byte[] buffer = new byte[BufferSize];

            TempFolder tempFolder = null;
            string sourceFilePath = null;
            FileStream sourceFileStream = null;
            string targetFilePath = null;
            FileStream targetFileStream = null;

            bool everythingOk = false;

            try
            {
                // 1) Read the conversion ID

                int conversionId;
                try
                {
                    conversionId = ReceiveInt();
                }
                catch (SocketException e)
                {
                    if (e.NativeErrorCode == 10060)
                    {
                        // This block handles the special case of an external service performing
                        // an health check on this server. The external service can just try to
                        // open a TCP connection, without sending anything. The simple fact that
                        // the server accepted the connection guarantees it is working.
                        // So if someone opens a TCP connection and doesn't send nothing, I fail
                        // gracefully. A socket exception with error code 10060 means socket
                        // timeout: I assume it is a health check.
                        healthCheck = true;
                        everythingOk = true;
                        return;
                    }
                    else
                    {
                        // In case of any other exception than socket timeout, throw the exception
                        // again to handle it regularly.
                        throw e;
                    }
                }
                log.Info("received conversion id: " + conversionId);


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

                tempFolder = new TempFolder(conversionId);
                log.Info("created temp folder " + tempFolder.ToString());

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
                everythingOk = true;
            }
            catch (Exception e)
            {
                log.Error("exception", e);
                StatusCodes statusCode = (e is ProtocolException ? ((ProtocolException)e).statusCode : StatusCodes.InternalServerError);
                string statusCodeDescription = "status code " + (int)statusCode + " (" + statusCode + ")";
                try
                {
                    SendInt((int)statusCode);
                    log.Info("sent " + statusCodeDescription);
                }
                catch (Exception ee)
                {
                    log.Warn("exception while sending " + statusCodeDescription, ee);
                }
            }
            finally
            {
                // Close streams
                if (sourceFileStream != null) sourceFileStream.Close();
                if (targetFileStream != null) targetFileStream.Close();
                
                // Delete temp folder if everything ok, or move it to errors repository
                if (tempFolder != null)
                {
                    tempFolder.Release(!everythingOk);
                }

                socket.Shutdown(SocketShutdown.Both);
                socket.Close();
                if (!healthCheck) log.Info("connection closed");
            }
        }

        private int ReceiveInt()
        {
            byte[] buffer = new byte[4];
            int bytesRead = socket.Receive(buffer, sizeof(int), 0);

            // Pay attention to endianess
            int networkValue = BitConverter.ToInt32(buffer, 0);
            int realValue = IPAddress.NetworkToHostOrder(networkValue);

            return realValue;
        }

        private void SendInt(int value)
        {
            // Pay attention to endianess
            int networkValue = IPAddress.HostToNetworkOrder(value);

            byte[] buffer = BitConverter.GetBytes(networkValue);
            int bytesSent = socket.Send(buffer, sizeof(int), 0);
        }
    }
}
