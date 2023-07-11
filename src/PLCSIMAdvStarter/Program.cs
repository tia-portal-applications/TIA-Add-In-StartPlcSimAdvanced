using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Microsoft.Win32;

namespace PLCSIMAdvStarter
{
    internal class Program
    {
        public const string c_PlcSimDllName = @"Siemens.Simatic.Simulation.Runtime.Api.x64";

        private static void Main(string[] args)
        {
            try
            {
#if DEBUG
            Debugger.Launch();
#endif
                /* Incase of args are null or empty start process
                 * as monitor process which observes PLCSIMAdvStarter.exe
                 * process and handles close operation            
                 */
                if (args == null || args.Length == 0)
                {
                    ProcessMonitor.MonitorPLCSIMAdvProcess();
                }
                else
                {
                    // Attach to AssemblyResolver event and start tia portal
                    var logfilepath = GetLogFilePath(args[2]);
                    AppDomain.CurrentDomain.AssemblyResolve += Resolver.LatestOpennessVersionAndPLCSimResolver;
                    Run.StartPortal(args, logfilepath);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Returns the path of the logfile
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static string GetLogFilePath(string pathToPlcSimFolder)
        {
            var path = Path.Combine(pathToPlcSimFolder, "PLSCIMAdv.log");

            if (!File.Exists(path))
            {
                var fileinfo = File.Create(path);
                fileinfo.Dispose();
            }

            return path;
        }
    }
}