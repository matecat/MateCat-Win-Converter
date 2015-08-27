using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace LegacyOfficeConverter
{

    public class Server
    {
        private const int SocketPollMicroseconds = 250000;

        private int port;
        private DirectoryInfo tmpDir;
        private int queueSize;
        private int convertersPoolSize;
        private string OCRConsolePath;

        private bool running = true;
        private bool stopped = false;

        public Server(int port, DirectoryInfo tmpDir, int queueSize, int convertersPoolSize, string OCRConsolePath)
        {
            this.port = port;
            this.tmpDir = tmpDir;
            this.queueSize = queueSize;
            this.convertersPoolSize = convertersPoolSize;
            this.OCRConsolePath = OCRConsolePath;
        }

        public void Start()
        {
            running = true;
            stopped = false;

            Object fileSystemLock = new Object();
            ConvertersRouter fileConverter = new ConvertersRouter(convertersPoolSize, OCRConsolePath);

            Socket server = null;
            EndPoint endPoint = new IPEndPoint(IPAddress.Any, port);
            try
            {
                server = new Socket(endPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
                server.Blocking = true;
                server.Bind(endPoint);
                server.Listen(queueSize);

                Console.WriteLine("Conversion server is ready to accept requests.");

                while (running)
                {
                    if (server.Poll(SocketPollMicroseconds, SelectMode.SelectRead))
                    {
                        Socket clientSocket = server.Accept();
                        Client connection = new Client(clientSocket, tmpDir, fileConverter, fileSystemLock);
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
                Thread.Sleep(500);
            }
        }
    }
}
