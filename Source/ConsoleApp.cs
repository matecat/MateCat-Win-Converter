using log4net;
using log4net.Config;
using System;
using System.Configuration;
using System.Net;
using System.Threading;
using Translated.MateCAT.WinConverter.Converters;
using static System.Reflection.MethodBase;

[assembly: XmlConfigurator(Watch = true)]

namespace Translated.MateCAT.WinConverter
{
    class ConsoleApp
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        static void Main(string[] args)
        {
            // If port is zero socket will attach on the first available port between 1024 and 5000 (see https://goo.gl/t4MBUr)
            int port = int.Parse(ConfigurationManager.AppSettings.Get("Port"));
            // The socket's queue size for incoming connections (see https://goo.gl/IIFY20)
            int queue = int.Parse(ConfigurationManager.AppSettings.Get("QueueSize"));
            int convertersPoolSize = int.Parse(ConfigurationManager.AppSettings.Get("ConvertersPoolSize"));

            // Greet the user and recap params
            Console.WriteLine("MateCAT WinConverter!");
            Console.WriteLine("Guessed external IP is: " + GuessLocalIPv4().ToString());
            Console.WriteLine("Press ESC to stop the server and quit");
            Console.WriteLine();

            log.Info("MateCAT WinConverter is starting");

            // Create the main conversion class
            IConverter converter = new ConvertersRouter(convertersPoolSize);

            // Then create and start the conversion server
            ConversionServer.ConversionServer server = new ConversionServer.ConversionServer(port, queue, converter);
            server.Start();

            // Press ESC to stop the server. 
            // If others keys are pressed, remember to the user that only ESC works.
            while (Console.ReadKey(true).Key != ConsoleKey.Escape)
            {
                Console.WriteLine("Press ESC to stop the server.");
            }

            // ESC key pressed, shutdown everything and say goodbye
            log.Info("User asked to stop");
            server.Stop();

            Console.WriteLine("Closing MateCAT WinConverter right now");
            // Let the user read the goodbye message
            Thread.Sleep(1000);
        }

        static IPAddress GuessLocalIPv4()
        {
            IPAddress[] localIPs = Dns.GetHostAddresses(Dns.GetHostName());
            foreach (IPAddress a in localIPs)
            {
                if (a.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork) return a;
            }
            return IPAddress.Loopback;
        }
    }
}
