using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using Siemens.Engineering;

namespace PLCSIMAdvStarter
{
    internal static class ProcessMonitor
    {
        private const string c_AddInMainProcess = "PLCSIMAdvStarter";

        /// <summary>
        /// Monitor the PLCSIM Advanced Process
        /// </summary>
        public static void MonitorPLCSIMAdvProcess()
        {
            var monitor_thread = new Thread(() =>
            {
                while (true)
                {
                    ///TODO ask why sleeping
                    Thread.Sleep(5000);

                    if (Monitor_PLCSIMAdvStarter() || Monitor_PLCSIMAdv()) ShutDownGracefully();
                }
            });

            monitor_thread.Start();
        }

        /// <summary>
        /// Monitor the the TIA-Portal process
        /// </summary>
        /// <param name="ProcessID">Id of the actual process</param>
        public static void MonitorTIAPortalProcess(int ProcessID)
        {
            var monitor_thread = new Thread(() =>
            {
                while (true)
                {
                    if (Monitor_TIA_Portal(ProcessID)) ShutDownGracefully();

                    Thread.Sleep(1000);
                }
            });

            monitor_thread.Start();
        }

        /// <summary>
        /// Shuts the process down
        /// </summary>
        private static void ShutDownGracefully()
        {
            var path = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            path = Path.Combine(path, "PLCSIMAdvAddIn");

            File.Delete(Path.Combine(path, "StateFile.txt"));

            Process.GetProcessesByName(c_AddInMainProcess).ToList().ForEach(process =>
            {
                if (Process.GetCurrentProcess().Id != process.Id)
                    process.Kill();
            });

            Process.GetCurrentProcess().Kill();
        }

        /// <summary>
        /// Monitors the TIA-Portal
        /// </summary>
        /// <param name="processID">Id of the actual process</param>
        /// <returns></returns>
        private static bool Monitor_TIA_Portal(int processID)
        {
            return TiaPortal.GetProcess(processID) == null;
        }

        /// <summary>
        /// Monitors the PLCSIM Advanced starter
        /// </summary>
        /// <returns></returns>
        private static bool Monitor_PLCSIMAdvStarter()
        {
            return Process.GetProcessesByName(c_AddInMainProcess).ToList().Count < 2;
        }

        /// <summary>
        /// Monitors the PLCSIM Advanced
        /// </summary>
        /// <returns></returns>
        private static bool Monitor_PLCSIMAdv()
        {
            return Process.GetProcessesByName("Siemens.Simatic.PlcSim.Advanced.UserInterface").Length == 0;
        }
    }
}