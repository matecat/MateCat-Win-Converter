using log4net;
using System;
using System.IO;
using System.Net.Sockets;
using static System.Reflection.MethodBase;
using Translated.MateCAT.WinConverter.Converters;
using Translated.MateCAT.WinConverter.Utils;
using System.Threading;

namespace Translated.MateCAT.WinConverter.ConversionServer
{
    public class ConversionRequest
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        public const int BufferSize = 10*1024*1024; // 10 MB
        public const int SocketTimeout = 5000; // in milliseconds

        private readonly Socket socket;
        private readonly IConverter fileConverter;
        private readonly ConversionServer.Counter serverConnectionsCounter;

        public ConversionRequest(Socket socket, IConverter fileConverter, ConversionServer.Counter serverConnections)
        {
            this.socket = socket;
            this.fileConverter = fileConverter;
            serverConnectionsCounter = serverConnections;
        }

        public void Run(object stateInfo)
        {
            if (!socket.Connected) throw new InvalidOperationException("Socket is not connected!");

            int totalTime = Environment.TickCount;

            int bytesRead, bytesSent;
            byte[] buffer = new byte[BufferSize];

            TempFolder tempFolder = null;
            string sourceFilePath = null;
            FileStream sourceFileStream = null;
            string targetFilePath = null;
            FileStream targetFileStream = null;

            bool healthCheck = false;
            bool everythingOk = false;

            try
            {
                // Set socket timeouts
                socket.ReceiveTimeout = SocketTimeout;
                socket.SendTimeout = SocketTimeout;

                // 1) Read the conversion ID
                int conversionId;
                try
                {
                    conversionId = ReceiveInt();
                }
                catch (InvalidOperationException)
                {
                    // If an external service performs a TCP health check on
                    // this on this service, it will shutdown the connection
                    // just after it is accepted by this server. In this catch
                    // I handle just this special case. Don't log anything,
                    // just return successfully.
                    healthCheck = true;
                    everythingOk = true;
                    return;
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
                    bytesRead = socket.Receive(buffer, Math.Min(BufferSize, fileSize - totalBytesRead), 0);
                    sourceFileStream.Write(buffer, 0, bytesRead);
                    totalBytesRead += bytesRead;
                }
                sourceFileStream.Close();
                log.Info("source file received");


                // 5) Convert the source file
                log.Info("converting file...");
                int conversionTime = Environment.TickCount;
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
                    throw new ProtocolException(StatusCodes.InternalServerError, e);
                }
                if (!converted)
                {
                    throw new ProtocolException(StatusCodes.UnsupportedConversion);
                }
                conversionTime = Environment.TickCount - conversionTime;
                log.Info("file converted");


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
                while ((bytesRead = targetFileStream.Read(buffer, 0, BufferSize)) > 0)
                {
                    bytesSent = socket.Send(buffer, bytesRead, 0);
                }
                log.Info("target file sent");
                targetFileStream.Close();

                // If execution arrives here, everything went well!
                everythingOk = true;

                // Log timings
                totalTime = Environment.TickCount - totalTime;
                log.Info("timings: total " + totalTime / 1000 + "s, conversion " + conversionTime / 1000 + "s ("+ (int)(((double)conversionTime / totalTime) * 100) +"%)");
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
                Interlocked.Decrement(ref serverConnectionsCounter.value);
            }
        }

        private int ReceiveInt()
        {
            const int bytesToRead = sizeof(int); // = 4
            int totalBytesRead = 0;
            byte[] buffer = new byte[bytesToRead];
            while (totalBytesRead < bytesToRead)
            {
                int bytesRead = socket.Receive(buffer, totalBytesRead, bytesToRead - totalBytesRead, 0);
                if (bytesRead == 0)
                {
                    // Flows enters here if the remote host interrupted the connection 
                    // (see https://goo.gl/sPrsZf)
                    throw new InvalidOperationException();
                }
                totalBytesRead += bytesRead;
            }

            // Pay attention to endianess: the standard for networking (and for Java) is Big Endian
            if (BitConverter.IsLittleEndian) Array.Reverse(buffer);

            return BitConverter.ToInt32(buffer, 0);
        }

        private void SendInt(int value)
        {
            byte[] buffer = BitConverter.GetBytes(value);
            // Pay attention to endianess: the standard for networking (and for Java) is Big Endian
            if (BitConverter.IsLittleEndian) Array.Reverse(buffer);

            int bytesSent = socket.Send(buffer);
        }
    }
}
