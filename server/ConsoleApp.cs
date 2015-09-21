using System;
using System.Configuration;
using System.Net;
using System.Threading;
using Translated.MateCAT.LegacyOfficeConverter.ConversionServer;
using Translated.MateCAT.LegacyOfficeConverter.Converters;

namespace Translated
{
    class ConsoleApp
    {
        static void Main(string[] args)
        {
            // If port is zero socket will attach on the first available port between 1024 and 5000 (see https://goo.gl/t4MBUr)
            int port = int.Parse(ConfigurationManager.AppSettings.Get("Port"));
            // The socket's queue size for incoming connections (see https://goo.gl/IIFY20)
            int queue = int.Parse(ConfigurationManager.AppSettings.Get("QueueSize"));
            int convertersPoolSize = int.Parse(ConfigurationManager.AppSettings.Get("ConvertersPoolSize"));

            // Greet the user and recap params
            Console.WriteLine("Hello sir.");
            Console.WriteLine();
            Console.WriteLine("LegacyOfficeConverter is starting.");
            Console.WriteLine();
            Console.WriteLine("Guessed external IP is: " + GuessLocalIPv4().ToString());
            Console.WriteLine();
            Console.WriteLine("Starting conversion server...");
            Console.WriteLine("Press ESC to stop the server and quit.");
            Console.WriteLine();

            // Create the main conversion class
            IConverter converter = new ConvertersRouter(convertersPoolSize);

            // Then create and start the conversion server
            ConversionServer server = new ConversionServer(port, queue, converter);
            server.Start();

            // Press ESC to stop the server. 
            // If others keys are pressed, remember to the user that only ESC works.
            while (Console.ReadKey(true).Key != ConsoleKey.Escape)
            {
                Console.WriteLine("Press ESC to stop the server.");
            }

            // ESC key pressed, shutdown everything and say goodbye
            Console.WriteLine();
            Console.WriteLine("ESC key pressed.");
            Console.WriteLine("Stopping the server...");
            server.Stop();

            Console.WriteLine("Server stopped.");
            Console.WriteLine("Good bye sir.");
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
