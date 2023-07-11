using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Connection;
using Siemens.Engineering.Download;
using Siemens.Engineering.Download.Configurations;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.Online;
using Siemens.Engineering.Online.Configurations;
using Siemens.Simatic.Simulation.Runtime;

namespace PLCSIMAdvStarter
{
    public static class Run
    {
        /// <summary>
        ///     Adds watcher to StateFile.txt files
        ///     and starts sim. for given PLC in args[]
        /// </summary>
        /// <param name="args">Arguments of the cli process</param>
        /// <param name="logfilepath">Path to the logfile</param>
        public static void StartPortal(string[] args, string logfilepath)
        {
#if DEBUG
            Debugger.Launch();
#endif
            // Fetch folder path of Resources ('PLCSIMAdvStarter.exe' and 'StateFile.txt')

            startStopSimulationState = start_State;
            var isProcessCancelled = false;

            m_logfilepath = logfilepath;
            m_pathToStateFile = args[2];

            var folderPath = m_pathToStateFile;

            try
            {
                // Fetch plcName and processID from passed arguments
                plcName = args[0];
                var processID = int.Parse(args[1]);

                ProcessMonitor.MonitorTIAPortalProcess(processID);

                // Attach to tia portal by given processID
                _tiaportal = TiaPortal.GetProcess(processID).Attach();
                if (IsPLCCompatible(plcName))
                {
                    // Start simulation of corresponding PLC instance
                    StartSimulation(out isProcessCancelled);

                    if (isProcessCancelled) return;

                    // Watch changes for 'StateFile.txt'
                    using (var watcher = new FileSystemWatcher(folderPath, "StateFile.txt"))
                    {
                        watcher.NotifyFilter = NotifyFilters.LastWrite;

                        watcher.EnableRaisingEvents = true;
                        // 'OnChanged' event trigerred when 'StateFile.txt' edited by AddIn
                        watcher.Changed += OnChanged;

                        // Continue process until all PLC instances stopped
                        while (isContinueListenFromStateFile) Thread.Sleep(60000);
                    }
                }
            }
            catch (Exception e)
            {
                File.AppendAllText(m_logfilepath,
                    $@"[{DateTime.Now}] ERROR : {startStopSimulationState} Simulation of {plcName} via PLCSIM Advanced Simulator failed : {e.Message}" +
                    Environment.NewLine);
            }
            finally
            {
                // Remove 'StateFile.txt' and end current process
                File.Delete(Path.Combine(folderPath, "StateFile.txt"));
                Process.GetCurrentProcess().Kill();
            }
        }

        #region Constants and Fields

        private const string c_PlcSimAdvancedRuntimeRegistry =
            @"SOFTWARE\Wow6432Node\Siemens\Shared Tools\PLCSIMADV_SimRT";

        private const string c_PlcSimAdvancedUserInterfaceExe = @"Siemens.Simatic.PlcSim.Advanced.UserInterface.exe";
        private const string c_MethodName_RegisterInstance = @"RegisterInstance";
        private const string c_MethodName_Shutdown = @"Shutdown";
        private const string start_State = "Start";
        private const string stop_State = "Stop";

        public static List<string> compatiblePLCDevices = new List<string>
        {
            "System:Device.S71500",
            "System:Device.ET200SP"
        };

        private static Assembly m_PlcSimAdvancedRuntimeManager;
        private static Process simProcess;
        private static bool isContinueListenFromStateFile = true;
        private static string startStopSimulationState;
        private static readonly Dictionary<string, object> m_PlcInstances = new Dictionary<string, object>();
        private static readonly Dictionary<string, DeviceItem> plcDevices = new Dictionary<string, DeviceItem>();
        private static TiaPortal _tiaportal;
        private static ExclusiveAccess exclusiveAccess;
        private static string plcName;
        public static string m_logfilepath;
        private static OnlineConfigurationDelegate m_OnlineConfigurationDelegate;
        public static string m_pathToStateFile;

        #endregion


        #region private methods

        /// <summary>
        /// Preconfiguration of the Download delegate
        /// </summary>
        /// <param name="preConfiguration">Contains the information about the download configuration</param>
        private static void PreConfigureDownloadDelegate(DownloadConfiguration preConfiguration)
        {
            switch (preConfiguration)
            {
                case CheckBeforeDownload check:
                    check.Checked = true;
                    return;
                case StopModules modules:
                    modules.CurrentSelection = StopModulesSelections.StopAll;
                    return;
                case OverwriteSystemData overwrite:
                    overwrite.CurrentSelection = OverwriteSystemDataSelections.Overwrite;
                    return;
                case ConsistentBlocksDownload consistent:
                    consistent.CurrentSelection = ConsistentBlocksDownloadSelections.ConsistentDownload;
                    return;
                case AlarmTextLibrariesDownload alarmtext:
                    alarmtext.CurrentSelection = AlarmTextLibrariesDownloadSelections.ConsistentDownload;
                    break;
            }
        }

        /// <summary>
        /// Postconfiguration of the Download delegate
        /// </summary>
        /// <param name="postConfiguration">Contains the information about the download configuration </param>
        private static void PostConfigureDownloadDelegate(DownloadConfiguration postConfiguration)
        {
            if (postConfiguration is StartModules modules)
                modules.CurrentSelection = StartModulesSelections.StartModule;
        }

        /// <summary>
        ///     Starts to download the PLC device
        /// </summary>
        /// <param name="provider">Object DownloadProvider</param>
        /// <param name="targerInterface">Object ConfigurationTargetInterface</param>
        /// <param name="isProcessCancelled">Bool to check if the download was cancelled</param>
        /// <returns></returns>
        private static bool DownloadPLCDevice(DownloadProvider provider, ConfigurationTargetInterface targerInterface,
            out bool isProcessCancelled)
        {
            isProcessCancelled = false;
            try
            {
                if (exclusiveAccess.IsCancellationRequested)
                {
                    isProcessCancelled = true;
                    return false;
                }

                var downloadResult = provider.Download(targerInterface, PreConfigureDownloadDelegate,
                    PostConfigureDownloadDelegate,
                    DownloadOptions.Hardware | DownloadOptions.Software);

                return downloadResult.State == DownloadResultState.Success;
            }
            catch (Exception ex)
            {
                File.AppendAllText(m_logfilepath,
                    $@"[{DateTime.Now}] ERROR : {plcName} could not be downloaded.Exception : {ex.Message}" +
                    Environment.NewLine);
                return false;
            }
        }

        /// <summary>
        ///     Starts sim for {plcName}
        /// </summary>
        /// <param name="isProcessCancelled">Bool to check if the download was cancelled</param>
        private static void StartSimulation(out bool isProcessCancelled)
        {
            isProcessCancelled = false;
            exclusiveAccess = _tiaportal.ExclusiveAccess("PLCSIM Advanced Simulation process started...");

            if (simProcess == null)
            {
                // Start process of PlcSimAdvanced
                var pathToProgramFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
                var pathToPlcSimAdvPath = Path.Combine(pathToProgramFilesX86, "Siemens\\Automation\\PLCSIMADV\\bin");
                var plcSimExePath = Path.Combine(pathToPlcSimAdvPath, c_PlcSimAdvancedUserInterfaceExe);
                var dirName = Path.GetDirectoryName(plcSimExePath);

                var plcSimProcessStartInfo = new ProcessStartInfo
                {
                    FileName = plcSimExePath,
                    WorkingDirectory = dirName
                };
                simProcess = Process.Start(plcSimProcessStartInfo);
            }

            // Load Assembly PlcSimAdvanced RuntimeManager
            m_PlcSimAdvancedRuntimeManager = LoadPLCSIMAssembly();

            // Register Instance
            exclusiveAccess.Text = $"Register and PowerOn {plcName} Instance...";


            m_PlcInstances.Add(plcName, RegisterInstance(plcName, m_PlcSimAdvancedRuntimeManager));

            // Set StoragePath
            var storagePath = m_PlcInstances[plcName].GetType().GetProperty(@"StoragePath");
            storagePath.GetSetMethod().Invoke(m_PlcInstances[plcName],
                new object[] { Path.Combine(Path.GetTempPath(), plcName) });

            // Power On Instance
            var powerOn = m_PlcInstances[plcName].GetType().GetMethod(@"PowerOn", new[] { typeof(uint) });
            powerOn?.Invoke(m_PlcInstances[plcName], new object[] { (uint)60000 });

            // Fetch PLC Device from parsed plc name
            ProjectBase project = GetProject();
            plcDevices.Add(plcName, FetchPLCDevice(plcName, project));

            // Get IP Configuration ('ipAddress', 'subnetMask', 'gatewayAddress') from the corresponding PLC Device
            var networkType = "PROFINET";
            var profinetInterface =
                plcDevices[plcName].DeviceItems.FirstOrDefault(di2 => di2.Name.Contains(networkType));
            var networkInterface = profinetInterface.GetService<NetworkInterface>();
            string ipAddress = "", subnetMask = "", gatewayAddress = "";
            if (networkInterface != null && networkInterface.InterfaceType == NetType.Ethernet &&
                profinetInterface.GetAttribute("Label").ToString() == "X1")
            {
                var node = networkInterface.Nodes.First();
                ipAddress = node.GetAttribute("Address").ToString();
                subnetMask = node.GetAttribute("SubnetMask").ToString();
                if ((bool)node.GetAttribute("UseRouter"))
                    gatewayAddress = node.GetAttribute("RouterAddress").ToString();
            }

            // Set IP Suite for the corresponding PLC instance on PLCSIM Advanced
            var setIpSuite = m_PlcInstances[plcName].GetType()
                .GetMethod(@"SetIPSuite", new[] { typeof(uint), typeof(SIPSuite4), typeof(bool) });
            var suite = new SIPSuite4(ipAddress, subnetMask, gatewayAddress);
            setIpSuite?.Invoke(m_PlcInstances[plcName], new object[] { (uint)0, suite, true });

            // All devices should be on offline mode to Download any other device
            foreach (var device in plcDevices.Values) device.GetService<OnlineProvider>().GoOffline();

            if (exclusiveAccess.IsCancellationRequested)
            {
                isProcessCancelled = true;
                CancelSimulation(plcName);


                foreach (var device in plcDevices) device.Value.GetService<OnlineProvider>().GoOnline();

                exclusiveAccess.Dispose();
                return;
            }

            // Download to device and go online for the PLC Device
            exclusiveAccess.Text = $"Downloading {plcName} and {plcName} is going Online...";

            var onlineProvider = plcDevices[plcName].GetService<OnlineProvider>();
            var configuration = onlineProvider.Configuration;
            try
            {
                // Register for OnlineLegitimation to trust via "TlsVerificationConfigurationSelection.Trusted"
                m_OnlineConfigurationDelegate = OnlineCallBackMethod;
                configuration.OnlineLegitimation += m_OnlineConfigurationDelegate;
            }
            catch (Exception e)
            {
                File.AppendAllText(m_logfilepath,
                    $@"[{DateTime.Now}] ERROR : Getting Online state of {plcName} is failed : {e.Message}" +
                    Environment.NewLine);
                CancelSimulation(plcName);
            }

            try
            {
                // Download to PLC Device
                var downloadProvider = plcDevices[plcName].GetService<DownloadProvider>();
                // Scanning for the matching network interface is skipped in this demonstration
                var targetConfiguration = downloadProvider.Configuration.Modes.Find("PN/IE").PcInterfaces
                    .Find("PLCSIM", 1).TargetInterfaces.Find("1 X1");

                // Download to PLC Device
                if (DownloadPLCDevice(downloadProvider, targetConfiguration, out isProcessCancelled))
                {
                    onlineProvider.Configuration.ApplyConfiguration(targetConfiguration);

                    foreach (var device in plcDevices)
                        try
                        {
                            device.Value.GetService<OnlineProvider>().GoOnline();
                        }
                        catch (Exception ex)
                        {
                            ShutDownPLC(device.Key);
                        }

                    if (exclusiveAccess.IsCancellationRequested)
                    {
                        isProcessCancelled = true;
                        CancelSimulation(plcName);

                        foreach (var device in plcDevices) device.Value.GetService<OnlineProvider>().GoOnline();
                        return;
                    }

                    File.AppendAllText(m_logfilepath,
                        $@"[{DateTime.Now}] SUCCESS : Start simulation of {plcName} via PLCSIM Advanced Simulaton is successful" +
                        Environment.NewLine);
                }
                else if (isProcessCancelled)
                {
                    CancelSimulation(plcName);

                    foreach (var device in plcDevices) device.Value.GetService<OnlineProvider>().GoOnline();
                }
                else
                {
                    ShutDownPLC(plcName);
                }
            }

            finally
            {
                configuration.OnlineLegitimation -= m_OnlineConfigurationDelegate;
                exclusiveAccess?.Dispose();
                ShowLogDialog(m_logfilepath);
            }
        }

        /// <summary>
        /// Get the active project. Either from the LocalSessions (Multiuser) or from a single project
        /// </summary>
        /// <returns></returns>
        private static ProjectBase GetProject()
        {
            ProjectBase project = null;
            if (_tiaportal.Projects.Count == 0)
            {
                project = _tiaportal.LocalSessions.FirstOrDefault()?.Project;
            }
            else if (_tiaportal.Projects.Count > 0)
            {
                project =
                    _tiaportal.Projects.FirstOrDefault();
            }

            return project;
        }

        /// <summary>
        ///     Turns the PLC off
        /// </summary>
        /// <param name="plcName">Contains the name of the PLC</param>
        private static void ShutDownPLC(string plcName)
        {
            // Power off the instance
            var powerOff = m_PlcInstances[plcName]?.GetType().GetMethod(@"PowerOff", new[] { typeof(uint) });
            powerOff?.Invoke(m_PlcInstances[plcName], new object[] { (uint)60000 });

            // Cleanup the storage path
            var cleanUpStoragePath = m_PlcInstances[plcName].GetType().GetMethod(@"CleanupStoragePath");
            cleanUpStoragePath?.Invoke(m_PlcInstances[plcName], new object[] { });

            // Unregister from the instance
            var unRegister = m_PlcInstances[plcName].GetType().GetMethod(@"UnregisterInstance");
            unRegister?.Invoke(m_PlcInstances[plcName], new object[] { });

            plcDevices.Remove(plcName);
            m_PlcInstances.Remove(plcName);

            if (!m_PlcInstances.Any() && !plcDevices.Any())
            {
                exclusiveAccess.Text = "PLCSIM Advanced Simulation process shutting down...";
                // Shutdown the PLCSIM Advanced
                var shutDown = m_PlcSimAdvancedRuntimeManager?.GetType()
                    .GetMethod(c_MethodName_Shutdown, new[] { typeof(void) });
                shutDown?.Invoke(m_PlcSimAdvancedRuntimeManager, new object[] { });

                // Kill the PLCSIM Advanced process
                if (simProcess != null)
                {
                    simProcess.Kill();
                    Thread.Sleep(5000);
                }

                simProcess = null;
                isContinueListenFromStateFile = false;
            }
        }

        /// <summary>
        ///     Stops the PLCSIM Advanced simulation
        /// </summary>
        /// <param name="plcName">Contains the name of the PLC</param>
        private static void CancelSimulation(string plcName)
        {
            var folderPath = m_pathToStateFile;
            var filePath = Path.Combine(folderPath, "StateFile.txt");

            // Go offline for the PLC Instance               
            var onlineProvider = plcDevices[plcName].GetService<OnlineProvider>();
            onlineProvider?.GoOffline();

            // Power off the instance
            var powerOff = m_PlcInstances[plcName]?.GetType().GetMethod(@"PowerOff", new[] { typeof(uint) });
            powerOff?.Invoke(m_PlcInstances[plcName], new object[] { (uint)60000 });

            // Cleanup the storage path
            var cleanUpStoragePath = m_PlcInstances[plcName].GetType().GetMethod(@"CleanupStoragePath");
            cleanUpStoragePath?.Invoke(m_PlcInstances[plcName], new object[] { });

            // Unregister from the instance
            var unRegister = m_PlcInstances[plcName].GetType().GetMethod(@"UnregisterInstance");
            unRegister?.Invoke(m_PlcInstances[plcName], new object[] { });

            plcDevices.Remove(plcName);
            m_PlcInstances.Remove(plcName);

            using (var outStream = new StreamWriter(File.Open(filePath, FileMode.Truncate,
                       FileAccess.Write,
                       FileShare.ReadWrite)))
            {
                outStream.Write("");
                outStream.Write($"Cancel,{plcName}");
                outStream.Close();
            }
        }

        /// <summary>
        ///     Stops sim for {plcName}
        /// </summary>
        /// <param name="isProcessCancelled">Bool to check if the download was cancelled</param>
        private static void StopSimulation(out bool isProcessCancelled)
        {
            isProcessCancelled = false;

            // Go offline for the PLC Instance
            exclusiveAccess = _tiaportal.ExclusiveAccess($"{plcName} is going Offline");
            var onlineProvider = plcDevices[plcName].GetService<OnlineProvider>();
            onlineProvider?.GoOffline();

            exclusiveAccess.Text = $"PowerOff and UnRegister from {plcName} Instance...";
            // Power off the instance
            var powerOff = m_PlcInstances[plcName]?.GetType().GetMethod(@"PowerOff", new[] { typeof(uint) });
            powerOff?.Invoke(m_PlcInstances[plcName], new object[] { (uint)60000 });

            // Cleanup the storage path
            var cleanUpStoragePath = m_PlcInstances[plcName].GetType().GetMethod(@"CleanupStoragePath");
            cleanUpStoragePath?.Invoke(m_PlcInstances[plcName], new object[] { });

            // Unregister from the instance
            var unRegister = m_PlcInstances[plcName].GetType().GetMethod(@"UnregisterInstance");
            unRegister?.Invoke(m_PlcInstances[plcName], new object[] { });

            plcDevices.Remove(plcName);
            m_PlcInstances.Remove(plcName);

            File.AppendAllText(m_logfilepath,
                $@"[{DateTime.Now}] SUCCESS : Stop simulation of {plcName} via PLCSIM Advanced Simulaton is successful" +
                Environment.NewLine);
            // Shutdown PLCSIM process if all PLC instances unregistered
            if (!m_PlcInstances.Any() && !plcDevices.Any())
            {
                exclusiveAccess.Text = "PLCSIM Advanced Simulation process shutting down...";
                Thread.Sleep(2000);

                if (exclusiveAccess.IsCancellationRequested)
                {
                    exclusiveAccess?.Dispose();
                    isProcessCancelled = true;
                    return;
                }

                // Shutdown the PLCSIM Advanced
                var shutDown = m_PlcSimAdvancedRuntimeManager?.GetType()
                    .GetMethod(c_MethodName_Shutdown, new[] { typeof(void) });
                shutDown?.Invoke(m_PlcSimAdvancedRuntimeManager, new object[] { });

                // Kill the PLCSIM Advanced process
                simProcess.Kill();
                exclusiveAccess?.Dispose();

                Thread.Sleep(5000);

                simProcess = null;
                isContinueListenFromStateFile = false;
            }

            if (exclusiveAccess.IsCancellationRequested)
            {
                exclusiveAccess?.Dispose();
                isProcessCancelled = true;
                return;
            }

            //   feedbackService.Log(NotificationIcon.Success, $"Stop simulation of {plcName} via PLCSIM Advanced Simulaton is successful");              
            exclusiveAccess?.Dispose();
        }

        /// <summary>
        /// Checks if the state has changed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            var folderPath = m_pathToStateFile;

            var reader = new StreamReader(File.Open(e.FullPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite));

            var stateArguments = reader.ReadToEnd().Split(',');
            reader.Close();

            // Permit only one thread to handle Start/Stop
            if (startStopSimulationState == stateArguments[0] && plcName == stateArguments[1])
                return;

            startStopSimulationState = stateArguments[0];
            plcName = stateArguments[1];

            if (startStopSimulationState == start_State)
            {
                if (IsPLCCompatible(plcName))
                {
                    StartSimulation(out var isProcessCancelled);
                    if (isProcessCancelled) return;
                }
                else
                {
                    File.AppendAllText(m_logfilepath,
                        $@"[{DateTime.Now}] ERROR : Simulation of {plcName} via PLCSIM Advanced Simulator is not supported : " +
                        Environment.NewLine);
                }
            }
            else if (startStopSimulationState == stop_State)
            {
                StopSimulation(out var isProcessCancelled);
                if (isProcessCancelled) return;
            }
        }

        /// <summary>
        /// Registers the instance
        /// </summary>
        /// <param name="plcName"></param>
        /// <param name="m_PlcSimAdvancedRuntimeManager"></param>
        /// <returns></returns>
        private static object RegisterInstance(string plcName, Assembly m_PlcSimAdvancedRuntimeManager)
        {
            // Register Instance
            var instanceName = plcName;
            var myType =
                m_PlcSimAdvancedRuntimeManager.GetType(@"Siemens.Simatic.Simulation.Runtime.SimulationRuntimeManager");
            var cpuType = m_PlcSimAdvancedRuntimeManager.GetType(@"Siemens.Simatic.Simulation.Runtime.ECPUType");
            var myMethod = myType.GetMethod(c_MethodName_RegisterInstance, new[] { cpuType, typeof(string) });
            var plcInstance = myMethod?.Invoke(null, new object[] { 0x000005DC, instanceName });
            return plcInstance;
        }

        /// <summary>
        ///     Loads the Assembly of PLCSIM Advanced
        /// </summary>
        /// <returns></returns>
        private static Assembly LoadPLCSIMAssembly()
        {
            //  Load Assembly PlcSimAdvanced RuntimeManager
            var registryKey = Registry.LocalMachine.OpenSubKey(c_PlcSimAdvancedRuntimeRegistry);
            var value = registryKey?.GetValue("Path");
            var strRuntimePath = value.ToString();
            strRuntimePath = Path.Combine(strRuntimePath, @"API\4.0");
            var plcSimAdvancedRuntimeManager = Path.Combine(strRuntimePath, Program.c_PlcSimDllName + ".dll");
            var m_PlcSimAdvancedRuntimeManager = Assembly.LoadFile(plcSimAdvancedRuntimeManager);
            return m_PlcSimAdvancedRuntimeManager;
        }

        /// <summary>
        /// Fetch the PLC Device 
        /// </summary>
        /// <param name="plcName">Name of the PLC</param>
        /// <param name="project">Actual project</param>
        /// <returns></returns>
        private static DeviceItem FetchPLCDevice(string plcName, ProjectBase project)
        {
            return project.Devices.SelectMany(device => device.DeviceItems)
                .FirstOrDefault(deviceItem => deviceItem.Name == plcName);
        }

        /// <summary>
        ///     Checks if the PLC is compatible for the simulation
        /// </summary>
        /// <param name="plcName"></param>
        /// <returns></returns>
        private static bool IsPLCCompatible(string plcName)
        {
            var project = GetProject();
            var identifier = FetchPLCDevice(plcName, project)?.Parent.GetAttribute("TypeIdentifier").ToString();


            if (compatiblePLCDevices.IndexOf(identifier) != -1)
                return true;

            File.AppendAllText(m_logfilepath,
                $@"[{DateTime.Now}] ERROR : Unsupported Controller Type {plcName} via PLCSIM Advanced Simulator failed." +
                Environment.NewLine);
            return false;
        }

        /// <summary>
        /// Callback if the process is done
        /// </summary>
        /// <param name="onlineConfiguration"></param>
        private static void OnlineCallBackMethod(OnlineConfiguration onlineConfiguration)
        {
            if (onlineConfiguration is TlsVerificationConfiguration verificationConfiguration)
                verificationConfiguration.CurrentSelection = TlsVerificationConfigurationSelection.Trusted;
        }

        /// <summary>
        /// Output if the process is finished
        /// </summary>
        /// <param name="logfilepath"></param>
        private static void ShowLogDialog(string logfilepath)
        {
            var outputForm = new OutputForm("PLCSIM Advanced AddIn", "Information",
                $"For more info please refer to this path: {logfilepath} ");
            outputForm.Size = new Size(300, 200);
            outputForm.ShowInTaskbar = true;
            outputForm.BringToFront();
            outputForm.ShowDialog();
        }

        #endregion
    }
}