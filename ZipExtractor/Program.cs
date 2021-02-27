using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NLog;

namespace ZipExtractor
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var configuration = new NLog.Config.LoggingConfiguration();
            var logfile = new NLog.Targets.FileTarget("logfile") { FileName = $"ZipExtractor_{DateTime.Now.Date:ddMMyyyy}.log" };
            var logconsole = new NLog.Targets.ConsoleTarget("logconsole");

            configuration.AddRule(NLog.LogLevel.Trace, NLog.LogLevel.Fatal, logfile);
            configuration.AddRule(NLog.LogLevel.Trace, NLog.LogLevel.Fatal, logconsole);
            LogManager.Configuration = configuration;
             
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormMain());
        }
    }
}
