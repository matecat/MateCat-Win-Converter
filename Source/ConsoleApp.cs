using log4net;
using log4net.Config;
using NCrontab;
using System;
using System.Configuration;
using System.Threading;
using System.Timers;
using Translated.MateCAT.WinConverter.Converters;
using static System.Reflection.MethodBase;

[assembly: XmlConfigurator(Watch = true)]

namespace Translated.MateCAT.WinConverter
{
    class ConsoleApp
    {
        private static readonly ILog log = LogManager.GetLogger(GetCurrentMethod().DeclaringType);

        private const int secondsBeforeRestart = 30;

        private static ConversionServer.ConversionServer server;

        static void Main(string[] args)
        {
            // If port is zero socket will attach on the first available port between 1024 and 5000 (see https://goo.gl/t4MBUr)
            int port = int.Parse(ConfigurationManager.AppSettings.Get("Port"));
            // The socket's queue size for incoming connections (see https://goo.gl/IIFY20)
            int queue = int.Parse(ConfigurationManager.AppSettings.Get("QueueSize"));
            int convertersPoolSize = int.Parse(ConfigurationManager.AppSettings.Get("ConvertersPoolSize"));
            string restartCronExpr = ConfigurationManager.AppSettings.Get("SystemRestartCron");

            // Greet the user and recap params
            Console.WriteLine("MateCAT WinConverter!");
            Console.WriteLine("Press ESC to stop the server and quit");
            Console.WriteLine();

            log.Info("MateCAT WinConverter is starting");

            // Set system restart timer, if configured
            if (!String.IsNullOrWhiteSpace(restartCronExpr))
            {
                CrontabSchedule restartCron = CrontabSchedule.Parse(restartCronExpr);
                DateTime restartDate = restartCron.GetNextOccurrence(DateTime.Now);

                System.Timers.Timer restartTimer = new System.Timers.Timer();
                restartTimer.Elapsed += new ElapsedEventHandler(SystemRestart);
                restartTimer.Interval = (int)(restartDate - DateTime.Now).TotalMilliseconds;
                restartTimer.AutoReset = false;
                restartTimer.Enabled = true;
                log.Info("Auto restart enabled: will restart at "+ restartDate.ToShortTimeString() + " of "+ restartDate.ToShortDateString());
            }

            // Create the main conversion class
            IConverter converter = new ConvertersRouter(convertersPoolSize);

            // Then create and start the conversion server
            server = new ConversionServer.ConversionServer(port, queue, converter);
            server.Start();

            // Press ESC to stop the server.
            // If others keys are pressed, remember to the user that only ESC works.
            while (Console.ReadKey(true).Key != ConsoleKey.Escape)
            {
                Console.WriteLine("Press ESC to stop the server");
            }

            // ESC key pressed, shutdown everything and say goodbye
            log.Info("User pressed ESC: stopping the server");
            server.Stop();

            Console.WriteLine("Closing MateCAT WinConverter right now");
            // Let the user read the goodbye message
            Thread.Sleep(1000);
        }

        private static void SystemRestart(object source, ElapsedEventArgs e)
        {
            log.Info("Restart timeout reached: stopping the server");
            server.Stop();

            System.Diagnostics.Process.Start("shutdown.exe", "-r -f -t " + secondsBeforeRestart);
            log.Info("System will restart in " + secondsBeforeRestart + " seconds; run \"shutdown -a\" to abort");

            int noticeDelay = 5;
            int i = secondsBeforeRestart;
            while (i > 0)
            {
                Thread.Sleep(noticeDelay * 1000);
                i -= noticeDelay;
                Console.WriteLine(i +" seconds before system restart (if not aborted)");
            }
        }
    }
}
