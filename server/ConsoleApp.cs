using NDesk.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading;

namespace LegacyOfficeConverter
{
    class ConsoleApp
    {
        static void Main(string[] args)
        {
            // Here are all the defaults for command line params
            string cache = Path.GetTempPath();
            string ocrConsolePath = @"C:\Program Files (x86)\OCR Console\OcrCon.exe";
            IPAddress addr = GuessLocalIPv4();
            // If port is zero socket will attach on the first available port between 1024 and 5000 (see https://goo.gl/t4MBUr)
            int port = 11000;
            // The socket's queue size for incoming connections (see https://goo.gl/IIFY20)
            int queue = 100;
            int convertersPoolSize = 3;
            bool help = false;

            // Setup command line params
            var p = new OptionSet() {
                { "c|cache=",  "{PATH} to the folder where incoming/converted file will be cached; default: " + cache,
                    v => cache = v },
                { "a|addr=",  "the listener {ADDRESS}; default: " + addr, 
                    v => addr = IPAddress.Parse(v) },
                { "p|port=",  "the listener {PORT}; default: " + port, 
                    (int v) => port = v },
                { "q|queue=",  "the listener queue {SIZE}; default: " + queue, 
                    (int v) => queue = v },
                { "o|ocr=",  "the ocr console path, if custom",
                    v => ocrConsolePath = v },
                { "m|parallelism=",  "the {NUMBER} of Office instances to launch at startup in order to support parallel conversions; default: " + convertersPoolSize,
                    (int v) => convertersPoolSize = v },
                { "h|help",  "show this message and exit", 
                    v => help = v != null },
            };

            // Parse command line params
            List<string> extra;
            try
            {
                extra = p.Parse(args);
            }
            catch (OptionException e)
            {
                Console.Write("LegacyOfficeConverter error: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try 'LegacyOfficeConverter --help' for more information.");
                return;
            }

            // Show help if user requested it
            if (help)
            {
                ShowHelp(p);
                return;
            }

            // Greet the user and recap params
            Console.WriteLine("Hello sir.");
            Console.WriteLine("Running LegacyOfficeConverter with options:");
            Console.WriteLine("  cache: " + cache);
            Console.WriteLine("  addr : " + addr);
            Console.WriteLine("  port : " + port);
            Console.WriteLine("  queue: " + queue);
            Console.WriteLine("  ocr console: " + ocrConsolePath);
            Console.WriteLine("  parallelism: " + convertersPoolSize);
            Console.WriteLine("Try 'LegacyOfficeConverter --help' for information on options.");
            Console.WriteLine();
            Console.WriteLine("Starting conversion server...");
            Console.WriteLine("Press ESC to stop the server.");
            Console.WriteLine();

            // Start the conversion server in another thread
            Server server = new Server(addr, port, new DirectoryInfo(cache), queue, convertersPoolSize, ocrConsolePath);
            Thread t = new Thread(new ThreadStart(server.Start));
            t.Start();

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

        static void ShowHelp (OptionSet p)
        {
            Console.WriteLine("Usage: LegacyOfficeConverter [OPTIONS]+");
            Console.WriteLine();
            Console.WriteLine("Options:");
            p.WriteOptionDescriptions(Console.Out);
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
