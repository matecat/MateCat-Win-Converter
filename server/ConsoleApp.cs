using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Threading;

namespace LegacyOfficeConverter
{
    class ConsoleApp
    {
        static void Main(string[] args)
        {
            string cache = ConfigurationManager.AppSettings.Get("CachePath");
            // If specified cache path is empty, use the system default temp dir
            if (cache == "")
            {
                cache = Path.GetTempPath();
            }

            // If port is zero socket will attach on the first available port between 1024 and 5000 (see https://goo.gl/t4MBUr)
            int port = Int32.Parse(ConfigurationManager.AppSettings.Get("Port"));
            // The socket's queue size for incoming connections (see https://goo.gl/IIFY20)
            int queue = Int32.Parse(ConfigurationManager.AppSettings.Get("QueueSize"));
            int convertersPoolSize = Int32.Parse(ConfigurationManager.AppSettings.Get("ConvertersPoolSize"));
            string ocrConsolePath = ConfigurationManager.AppSettings.Get("OCRConsolePath");

            // Greet the user and recap params
            Console.WriteLine("Hello sir.");
            Console.WriteLine();
            Console.WriteLine("Running LegacyOfficeConverter with options:");
            Console.WriteLine("  cache: " + cache);
            Console.WriteLine("  port : " + port);
            Console.WriteLine("  queue: " + queue);
            Console.WriteLine("  ocr console: " + ocrConsolePath);
            Console.WriteLine("  parallelism: " + convertersPoolSize);
            Console.WriteLine("Options are stored in the .config file in the same dir of this exe.");
            Console.WriteLine();
            Console.WriteLine("Guessed external IP is: " + GuessLocalIPv4().ToString());
            Console.WriteLine();
            Console.WriteLine("Starting conversion server...");
            Console.WriteLine("Press ESC to stop the server and quit.");
            Console.WriteLine();

            // Start the conversion server in another thread
            ConversionServer server = new ConversionServer(port, new DirectoryInfo(cache), queue, convertersPoolSize, ocrConsolePath);
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
