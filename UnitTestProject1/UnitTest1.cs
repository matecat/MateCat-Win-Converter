using Microsoft.VisualStudio.TestTools.UnitTesting;
using LegacyOfficeConverter;
using System.IO;
using System.Threading;
using System.Net.Sockets;
using System.Net;
using System;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {

        Socket socket;

        [TestMethod]
        [DeploymentItem("test.doc")]
        public void TestMethod1()
        {
            int port = 11000;

            Server server = new Server(port, new DirectoryInfo(Path.GetTempPath()), 100, 1, null);
            Thread t = new Thread(new ThreadStart(server.Start));
            t.Start();

            while (!server.IsRunning())
            {
                Thread.Sleep(250);
            }

            EndPoint endPoint = new IPEndPoint(IPAddress.Loopback, port);

            socket = new Socket(endPoint.AddressFamily,
                SocketType.Stream,
                ProtocolType.Tcp);
            socket.Connect(endPoint);

            SendInt((int) FileTypes.doc);
            SendInt((int) FileTypes.docx);
            SendInt((int) new FileInfo("test.doc").Length);
            socket.SendFile("test.doc");
            int statudCode = ReceiveInt();
            int convertedSize = ReceiveInt();

            var buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = socket.Receive(buffer, buffer.Length, 0)) > 0)
            {
                // hello
            }
            socket.Close();

            Console.WriteLine((int) new FileInfo("test.doc").Length);

            //server.Stop();
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
