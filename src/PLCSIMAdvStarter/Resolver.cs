using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;

namespace PLCSIMAdvStarter
{
    public static class Resolver
    {
        public const string c_PlcSimDllName = @"Siemens.Simatic.Simulation.Runtime.Api.x64";
        public static Assembly LatestOpennessVersionAndPLCSimResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf(',');
            RegistryKey highestOpennessEntry = null;
            string fullPath = null;
            if (index == -1)
            {
                return null;
            }
            string name = args.Name.Substring(0, index);
            RegistryKey key = null;

            if (name == "Siemens.Engineering")
            {
                key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Siemens\\Automation\\Openness");
                var highestTiaEntry = key.OpenSubKey(GetHighestVersionName(key) + "\\PublicAPI");
                highestOpennessEntry = highestTiaEntry.OpenSubKey(GetHighestVersionName(highestTiaEntry));

                if (highestOpennessEntry == null)
                    return null;
                object oRegKeyValue = highestOpennessEntry.GetValue(name);
                if (oRegKeyValue != null)
                {
                    string filePath = oRegKeyValue.ToString();
                    fullPath = Path.GetFullPath(filePath);
                }
            }

            if (name == c_PlcSimDllName)
            {
                key =
                    Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Wow6432Node\Siemens\Shared Tools\PLCSIMADV_SimRT");
                var mainDirectory = key.GetValue("Path").ToString();
                var parentDirectory = Path.Combine(mainDirectory, @"API\4.0");
                fullPath = Path.GetFullPath(Path.Combine(parentDirectory, c_PlcSimDllName + ".dll"));
            }

            return File.Exists(fullPath) ? Assembly.LoadFrom(fullPath) : null;
        }

        /// <summary>
        /// Get subkeys inside registry entry that represent a version and return the highest (latest) key 
        /// </summary>
        /// <param name="root">Registry entry containing version keys</param>
        /// <returns>Subkey with highest version</returns>
        private static string GetHighestVersionName(RegistryKey root)
        {
            var subKeys = root.GetSubKeyNames();

            var TiaVersions = subKeys.Select(key => (Key: key, Versioned: new Version(key)));
            var highest = TiaVersions.Max(entry => entry.Versioned);
            return TiaVersions.First(v => v.Versioned == highest).Key;
        }
    }
}

