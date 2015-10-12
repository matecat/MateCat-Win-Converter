using log4net;
using System;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using Translated.MateCAT.WinConverter.Converters;
using static System.Reflection.MethodBase;

namespace Translated.MateCAT.WinConverter.ConversionServer
{
    public class ConversionServer
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

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
            // Start the server in another thread
            Thread t = new Thread(new ThreadStart(StartListening));
            t.Start();

            // Wait the server is ready before returning to the caller
            while (!running)
            {
                Thread.Sleep(200);
            }
        }

        private void StartListening()
        {
            Socket serverSocket = null;
            EndPoint endPoint = new IPEndPoint(IPAddress.Any, port);
            try
            {
                serverSocket = new Socket(endPoint.AddressFamily, SocketType.Stream, ProtocolType.Tcp);
                serverSocket.Blocking = true;
                serverSocket.Bind(endPoint);
                serverSocket.Listen(queueSize);

                log.Info("Server started, listening on port " + port);

                running = true;
                stopped = false;

                while (running)
                {
                    // Calling socket.Accept blocks the thread until the next incoming connection,
                    // making difficult to stop the server from another thread.
                    // The Poll always returns after the specified delay elapsed, or immeidately
                    // returns if it detects an incoming connection. It's the perfect method
                    // to make this loop regularly che the running var, ending gracefully 
                    // if requested.
                    if (serverSocket.Poll(SocketPollMicroseconds, SelectMode.SelectRead))
                    {
                        Socket clientSocket = serverSocket.Accept();
                        ConversionRequest connection = new ConversionRequest(clientSocket, converter);
                        Thread clientThread = new Thread(new ThreadStart(connection.Run));
                        clientThread.Start();
                    }
                }
            }
            catch (Exception e)
            {
                log.Error("Exception", e);
            }
            finally
            {
                if (serverSocket != null) serverSocket.Close();
                running = false;
                stopped = true;
                log.Info("Server stopped");
            }
        }

        public void Stop()
        {
            running = false;

            // Wait the server is stopped before returning to the caller
            while (!stopped)
            {
                Thread.Sleep(200);
            }
        }
    }
}
