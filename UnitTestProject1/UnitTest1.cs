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

        static int testPort = 11000;
        static Server testServer;
        static EndPoint testEndPoint = new IPEndPoint(IPAddress.Loopback, testPort);

        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {     
            testServer = new Server(testPort, new DirectoryInfo(Path.GetTempPath()), 100, 1, null);
            Thread t = new Thread(new ThreadStart(testServer.Start));
            t.Start();

            while (!testServer.IsRunning())
            {
                Thread.Sleep(250);
            }
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            testServer.Stop();
        }


        [TestMethod]
        [DeploymentItem("test.doc")]
        public void TestRegularFlow()
        {
            Socket socket;
            socket = new Socket(testEndPoint.AddressFamily,
                SocketType.Stream,
                ProtocolType.Tcp);
            socket.Connect(testEndPoint);

            SendInt(socket, (int)FileTypes.doc);
            SendInt(socket, (int) FileTypes.docx);
            SendInt(socket, (int) new FileInfo("test.doc").Length);
            socket.SendFile("test.doc");
            int statudCode = ReceiveInt(socket);
            int convertedSize = ReceiveInt(socket);

            var buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = socket.Receive(buffer, buffer.Length, 0)) > 0)
            {
                // hello
            }
            socket.Close();

            Console.WriteLine((int) new FileInfo("test.doc").Length);
        }

        [TestMethod]
        public void TestBadSourceFileType()
        {
            int statusCode = RequestConversion(int.MaxValue, (int) FileTypes.docx, 1, new byte[] { 0 });
            Assert.AreEqual((int)Errors.BadFileType, statusCode);
        }

        [TestMethod]
        public void TestBadTargetFileType()
        {
            int statusCode = RequestConversion((int)FileTypes.doc, int.MaxValue, 1, new byte[] { 0 });
            Assert.AreEqual((int)Errors.BadFileType, statusCode);
        }

        [TestMethod]
        public void TestBadFileSize()
        {
            int statusCode = RequestConversion((int)FileTypes.doc, (int)FileTypes.docx, 0, new byte[] { 0 });
            Assert.AreEqual((int)Errors.BadFileSize, statusCode);

            statusCode = RequestConversion((int)FileTypes.doc, (int)FileTypes.docx, -1, new byte[] { 0 });
            Assert.AreEqual((int)Errors.BadFileSize, statusCode);

            statusCode = RequestConversion((int)FileTypes.doc, (int)FileTypes.docx, int.MinValue, new byte[] { 0 });
            Assert.AreEqual((int)Errors.BadFileSize, statusCode);
        }

        private int ReceiveInt(Socket socket)
        {
            byte[] buffer = new byte[4];
            int bytesRead = socket.Receive(buffer, sizeof(int), 0);
            // The "network" value
            int netValue = BitConverter.ToInt32(buffer, 0);
            // The "windows" value
            int winValue = IPAddress.NetworkToHostOrder(netValue);
            return winValue;
        }

        private void SendInt(Socket socket, int value)
        {
            int netValue = IPAddress.HostToNetworkOrder(value);
            byte[] buffer = BitConverter.GetBytes(netValue);
            int bytesSent = socket.Send(buffer, sizeof(int), 0);
        }


        private int RequestConversion(int sourceFileType, int targetFileType, int fileSize, byte[] file)
        {
            Socket socket = null;
            try
            {
                socket = new Socket(testEndPoint.AddressFamily,
                    SocketType.Stream,
                    ProtocolType.Tcp);
                socket.Connect(testEndPoint);

                SendInt(socket, sourceFileType);
                SendInt(socket, targetFileType);
                SendInt(socket, fileSize);
                socket.Send(file, file.Length, 0);

                int statusCode = ReceiveInt(socket);

                if (statusCode != 0)
                {
                    return statusCode;
                }

                int convertedSize = ReceiveInt(socket);

                var buffer = new byte[1024];
                int bytesRead, totalBytesRead = 0;
                while ((bytesRead = socket.Receive(buffer, buffer.Length, 0)) > 0)
                {
                    totalBytesRead += bytesRead;
                }

                if (totalBytesRead != convertedSize)
                {
                    throw new Exception("Expected file length: " + convertedSize + " bytes; Received file length: " + totalBytesRead + " bytes");
                }

                return statusCode;
            }
            finally
            {
                if (socket != null) socket.Close();
            }
        }
    }
}
