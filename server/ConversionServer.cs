using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;

namespace LegacyOfficeConverter
{

    public class ConversionServer
    {
        private const int SocketPollMicroseconds = 250000;

        private int port;
        private DirectoryInfo tmpDir;
        private int queueSize;
        private int convertersPoolSize;
        private string OCRConsolePath;

        private bool running = false;
        private bool stopped = true;

        public ConversionServer(int port, DirectoryInfo tmpDir, int queueSize, int convertersPoolSize, string OCRConsolePath)
        {
            this.port = port;
            this.tmpDir = tmpDir;
            this.queueSize = queueSize;
            this.convertersPoolSize = convertersPoolSize;
            this.OCRConsolePath = OCRConsolePath;
        }

        public void Start()
        {
            Thread t = new Thread(new ThreadStart(StartListening));
            t.Start();
            while (!running)
            {
                // Wait the server is ready before returning to the caller
                Thread.Sleep(200);
            }
        }

        private void StartListening()
        {
            object fileSystemLock = new object();
            ConvertersRouter fileConverter = new ConvertersRouter(convertersPoolSize, OCRConsolePath);

            Socket server = null;
            EndPoint endPoint = new IPEndPoint(IPAddress.Any, port);
            try
            {
                server = new Socket(endPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
                server.Blocking = true;
                server.Bind(endPoint);
                server.Listen(queueSize);

                running = true;
                stopped = false;

                Console.WriteLine("Conversion server is ready to accept requests.");

                while (running)
                {
                    if (server.Poll(SocketPollMicroseconds, SelectMode.SelectRead))
                    {
                        Socket clientSocket = server.Accept();
                        ConversionRequest connection = new ConversionRequest(clientSocket, tmpDir, fileConverter, fileSystemLock);
                        Thread clientThread = new Thread(new ThreadStart(connection.Run));
                        clientThread.Start();
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            finally
            {
                running = false;
                stopped = true;
                if (server != null) server.Close();
                fileConverter.Dispose();
            }
        }

        public void Stop()
        {
            running = false;
            while (!stopped)
            {
                Thread.Sleep(200);
            }
        }
    }
}
