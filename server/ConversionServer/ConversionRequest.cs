using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using Translated.MateCAT.LegacyOfficeConverter.Converters;
using Translated.MateCAT.LegacyOfficeConverter.Utils;

namespace Translated.MateCAT.LegacyOfficeConverter.ConversionServer
{
    public class ConversionRequest
    {
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

            Console.WriteLine(socket.RemoteEndPoint + " connected");

            TempFolder tempFolder = null;
            string sourceFilePath = null;
            FileStream sourceFileStream = null;
            string targetFilePath = null;
            FileStream targetFileStream = null;

            try
            {
                // 1) Read the source file type

                int fileTypeValue = ReceiveInt();
                FileTypes sourceFileType = (FileTypes)fileTypeValue;
                Console.WriteLine(socket.RemoteEndPoint + " received input file type: " + fileTypeValue + " ("+ sourceFileType +")");
                if (!FileTypes.IsDefined(typeof(FileTypes), sourceFileType))
                {
                    throw new ProtocolException(StatusCodes.BadFileType);
                }


                // 2) Read the target file type

                int outputFileTypeValue = ReceiveInt();
                FileTypes targetFileType = (FileTypes)outputFileTypeValue;
                Console.WriteLine(socket.RemoteEndPoint + " received output file type: " + outputFileTypeValue + " (" + targetFileType + ")");
                if (!FileTypes.IsDefined(typeof(FileTypes), targetFileType))
                {
                    throw new ProtocolException(StatusCodes.BadFileType);
                }


                // 3) Read the source file size

                int fileSize = ReceiveInt();
                Console.WriteLine(socket.RemoteEndPoint + " received file size: " + fileSize);
                if (fileSize <= 0)
                {
                    throw new ProtocolException(StatusCodes.BadFileSize);
                }

                // 4) Read the source file and cache it on disk

                tempFolder = new TempFolder();
                sourceFilePath = tempFolder.getFilePath("source." + sourceFileType);
                sourceFileStream = new FileStream(sourceFilePath, FileMode.Create);
                int totalBytesRead = 0;
                Console.WriteLine(socket.RemoteEndPoint + " receiving file");
                while (totalBytesRead < fileSize)
                {
                    bytesRead = socket.Receive(buffer, BufferSize, 0);
                    sourceFileStream.Write(buffer, 0, bytesRead);
                    totalBytesRead += bytesRead;
                }
                sourceFileStream.Close();
                Console.WriteLine(socket.RemoteEndPoint + " received file");


                // 5) Convert the source file
                Console.WriteLine(socket.RemoteEndPoint + " converting file");
                bool converted = false;
                try
                {
                    targetFilePath = tempFolder.getFilePath("target." + targetFileType);
                    converted = fileConverter.Convert(sourceFilePath, (int)sourceFileType, targetFilePath, (int)targetFileType);
                    Console.WriteLine(socket.RemoteEndPoint + " converted file");
                }
                catch (Exception e)
                {
                    throw new ProtocolException(StatusCodes.BrokenFile, e);
                }
                if (!converted)
                {
                    throw new ProtocolException(StatusCodes.UnsupportedConversion);
                }


                // 6) Send back the ok status code

                SendInt((int)StatusCodes.Ok);
                Console.WriteLine(socket.RemoteEndPoint + " sent status code: 0");


                // 7) Send back the converted file size

                int filesize = 0;
                try
                {
                    filesize = (int)new System.IO.FileInfo(targetFilePath).Length;
                }
                catch (OverflowException e)
                {
                    throw new ProtocolException(StatusCodes.ConvertedFileTooBig, e);
                }
                SendInt(filesize);
                Console.WriteLine(socket.RemoteEndPoint + " sent converted file size: " + filesize);


                // 8) Send back the converted file

                buffer = new byte[BufferSize];
                targetFileStream = new FileStream(targetFilePath, FileMode.Open);
                Console.WriteLine(socket.RemoteEndPoint + " sending file");
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
                Console.WriteLine(socket.RemoteEndPoint + " sent file");
                targetFileStream.Close();

                // If execution arrives here, everything went well!
            }
            catch (ProtocolException e)
            {
                try
                {
                    Console.WriteLine(socket.RemoteEndPoint + " error: " + e);
                    SendInt((int)e.ErrorCode);
                    Console.WriteLine(socket.RemoteEndPoint + " sent error");
                }
                catch { }
            }
            catch (Exception e)
            {
                Console.WriteLine(socket.RemoteEndPoint + "error\n" + e);
                try
                {
                    SendInt((int)StatusCodes.InternalServerError);
                    Console.WriteLine(socket.RemoteEndPoint + " sent error");
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
                Console.WriteLine(remoteEndPoint + " closed connection");
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
