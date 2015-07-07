using NDesk.Options;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    class ConsoleApp
    {
        static void Main(string[] args)
        {
            FormatConverterPool<WordConverter> w = new FormatConverterPool<WordConverter>(10);
            (new Thread(() => {
                int i = 1;
                Console.WriteLine(i + ": start");
                w.Convert(@"C:\Users\Giuseppe\Desktop\doc test\"+ i + ".doc");
                Console.WriteLine(i + ": end");
            })).Start();
            (new Thread(() => {
                int i = 2;
                Console.WriteLine(i + ": start");
                w.Convert(@"C:\Users\Giuseppe\Desktop\doc test\" + i + ".doc");
                Console.WriteLine(i + ": end");
            })).Start();
            (new Thread(() => {
                int i = 3;
                Console.WriteLine(i + ": start");
                w.Convert(@"C:\Users\Giuseppe\Desktop\doc test\" + i + ".doc");
                Console.WriteLine(i + ": end");
            })).Start();
            (new Thread(() => {
                int i = 4;
                Console.WriteLine(i + ": start");
                w.Convert(@"C:\Users\Giuseppe\Desktop\doc test\" + i + ".doc");
                Console.WriteLine(i + ": end");
            })).Start();
            (new Thread(() => {
                int i = 5;
                Console.WriteLine(i + ": start");
                w.Convert(@"C:\Users\Giuseppe\Desktop\doc test\" + i + ".doc");
                Console.WriteLine(i + ": end");
            })).Start();
            Thread.Sleep(10000);
            w.Dispose();
            return;

            String cache = Path.GetTempPath();
            IPAddress addr = GuessLocalIPv4();
            // If port is zero socket will attach on the first available port between 1024 and 5000
            // (see https://goo.gl/t4MBUr)
            int port = 11000;
            // The socket's queue size for incoming connections (see https://goo.gl/IIFY20)
            int queue = 100;
            bool help = false;
            string ocrPath = @"C:\Program Files (x86)\OCR Console\OcrCon.exe";;

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
                    v => ocrPath = v },
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

            // Ok, start the real server
            Console.WriteLine("Hello sir.");
            Console.WriteLine("Running LegacyOfficeConverter with options:");
            Console.WriteLine("  cache: " + cache);
            Console.WriteLine("  addr : " + addr);
            Console.WriteLine("  port : " + port);
            Console.WriteLine("  queue: " + queue);
            Console.WriteLine("  ocr console: " + @ocrPath);
            Console.WriteLine("Try 'LegacyOfficeConverter --help' for information on options.");
            Console.WriteLine();
            Console.WriteLine("Starting conversion server...");
            Console.WriteLine("Press ESC to stop the server.");
            Console.WriteLine();

            // Start the conversion server in another thread
            Server soc = new Server(addr, port, new DirectoryInfo(cache), queue, ocrPath);
            Thread t = new Thread(new ThreadStart(soc.Start));
            t.Start();

            // Press any key to stop the server
            while (Console.ReadKey(true).Key != ConsoleKey.Escape)
            {
                Console.WriteLine("Press ESC to stop the server.");
            }

            Console.WriteLine();
            Console.WriteLine("ESC key pressed.");
            Console.WriteLine("Stopping the server...");
            soc.Stop();
            Console.WriteLine("Server stopped.");
            Console.WriteLine("Good bye sir.");
            Thread.Sleep(1000);
            return;
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
