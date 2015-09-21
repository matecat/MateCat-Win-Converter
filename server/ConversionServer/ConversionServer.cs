using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using Translated.MateCAT.LegacyOfficeConverter.Converters;

namespace Translated.MateCAT.LegacyOfficeConverter.ConversionServer
{

    public class ConversionServer
    {
        private const int SocketPollMicroseconds = 250000;

        private int port;
        private int queueSize;
        private IConverter converter;

        private bool running = false;
        private bool stopped = true;

        public ConversionServer(int port, int queueSize, IConverter converter)
        {
            this.port = port;
            this.queueSize = queueSize;
            this.converter = converter;
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
                        ConversionRequest connection = new ConversionRequest(clientSocket, converter);
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
