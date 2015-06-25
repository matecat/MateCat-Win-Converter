using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    public class Client
    {
        public const int BufferSize = 8192;

        private Socket socket;
        private DirectoryInfo tmpDir;
        private FileConverter fileConverter;
        private Object fileSystemLock;

        public Client(Socket socket, DirectoryInfo tmpDir, FileConverter fileConverter, Object fileSystemLock)
        {
            this.socket = socket;
            this.tmpDir = tmpDir;
            this.fileConverter = fileConverter;
            this.fileSystemLock = fileSystemLock;
        }

        public void Run()
        {
            int bytesRead, bytesSent;
            byte[] buffer = new byte[BufferSize];

            Console.WriteLine(socket.RemoteEndPoint + " connected");

            string tmpFilePath = null;
            FileStream tmpFileStream = null;
            string convertedFilePath = null;
            FileStream convertedFileStream = null;

            try
            {
                // 1) Read the input file type

                int fileTypeValue = ReceiveInt();
                FileTypes fileType = (FileTypes)fileTypeValue;
                Console.WriteLine(socket.RemoteEndPoint + " received file type: " + fileTypeValue + " ("+ fileType +")");
                if (!FileTypes.IsDefined(typeof(FileTypes), fileType))
                {
                    throw new ProtocolException(Errors.BadFileType);
                }


                // 2) Read the input file size

                int fileSize = ReceiveInt();
                Console.WriteLine(socket.RemoteEndPoint + " received file size: " + fileSize);
                if (fileSize <= 0)
                {
                    throw new ProtocolException(Errors.BadFileSize);
                }

                // 3) Read the input file and cache it on disk

                // Very rare, but two concurrent threads can try to write on the same file:
                // avoid it locking the FileStream opening block
                lock (fileSystemLock)
                {
                    do
                    {
                        tmpFilePath = tmpDir + Path.GetRandomFileName().Substring(0,8) + "." + fileType;
                    } while (File.Exists(tmpFilePath));
                    tmpFileStream = new FileStream(tmpFilePath, FileMode.Create);
                }
                int totalBytesRead = 0;
                Console.WriteLine(socket.RemoteEndPoint + " receiving file");
                while (totalBytesRead < fileSize)
                {
                    bytesRead = socket.Receive(buffer, BufferSize, 0);
                    tmpFileStream.Write(buffer, 0, bytesRead);
                    totalBytesRead += bytesRead;
                }
                tmpFileStream.Close();
                Console.WriteLine(socket.RemoteEndPoint + " received file");


                // 4) Convert the source file

                Console.WriteLine(socket.RemoteEndPoint + " converting file");
                try
                {
                    convertedFilePath = fileConverter.Convert(tmpFilePath);
                    Console.WriteLine(socket.RemoteEndPoint + " converted file");
                }
                catch (Exception e)
                {
                    throw new ProtocolException(Errors.BrokenFile, e);
                }


                // 5) Send back the ok status code

                SendInt(0);
                Console.WriteLine(socket.RemoteEndPoint + " sent status code: 0");


                // 6) Send back the converted file size

                int filesize = 0;
                try
                {
                    filesize = (int)new System.IO.FileInfo(convertedFilePath).Length;
                }
                catch (OverflowException e)
                {
                    throw new ProtocolException(Errors.ConvertedFileTooBig, e);
                }
                SendInt(filesize);
                Console.WriteLine(socket.RemoteEndPoint + " sent converted file size: " + filesize);


                // 7) Send back the converted file and then delete it

                buffer = new byte[BufferSize];
                convertedFileStream = new FileStream(convertedFilePath, FileMode.Open);
                Console.WriteLine(socket.RemoteEndPoint + " sending file");
                while (true)
                {
                    bytesRead = convertedFileStream.Read(buffer, 0, BufferSize);
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
                convertedFileStream.Close();

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
                    SendInt((int)Errors.InternalServerError);
                    Console.WriteLine(socket.RemoteEndPoint + " sent error");
                }
                catch { }
            }
            finally
            {
                // Close streams                
                if (tmpFileStream != null) tmpFileStream.Close();
                if (convertedFileStream != null) convertedFileStream.Close();
                
                // Delete temp files
                if (tmpFilePath != null)
                {
                    try
                    {
                        File.Delete(tmpFilePath);
                    }
                    catch { }
                }
                if (convertedFilePath != null)
                {
                    try
                    {
                        File.Delete(convertedFilePath);
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
