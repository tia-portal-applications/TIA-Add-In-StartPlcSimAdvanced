using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Permissions;
using System.Windows.Forms;
using PLCSIMAdv.Properties;
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.AddIn.Permissions;
using Siemens.Engineering.HW;
using Siemens.Engineering.Online;
using Process = Siemens.Engineering.AddIn.Utilities.Process;

namespace PLCSIMAdvAddIn
{
    public class AddIn : ContextMenuAddIn
    {
        private const string c_StartState = "Start";
        private const string c_StopState = "Stop";


        //Process process;
        /// <summary>
        ///     The display name of the Add-In.
        /// </summary>
        private const string s_DisplayNameOfAddIn = "PLCSIM Advanced Simulation";

        /// <summary>
        ///     The global TIA Portal Object
        ///     <para>It will be used in the TIA Add-In.</para>
        /// </summary>
        private readonly TiaPortal _tiaportal;

        /// <summary>
        ///     The constructor of the AddIn.
        ///     Creates an object of the class AddIn
        ///     Called from AddInProvider, when the first
        ///     right-click is performed in TIA
        ///     Motherclass' constructor of ContextMenuAddin
        ///     will be executed, too.
        /// </summary>
        /// <param name="tiaportal">
        ///     Represents the actual used TIA Portal process.
        /// </param>
        public AddIn(TiaPortal tiaportal) : base(s_DisplayNameOfAddIn)
        {
            /*
            * The acutal TIA Portal process is saved in the
            * global TIA Portal variable _tiaportal
            * tiaportal comes as input Parameter from the
            * AddInProvider
            */
            _tiaportal = tiaportal;
        }

        /// <summary>
        ///     The method is supplemented to include the Add-In
        ///     in the Context Menu of TIA Portal.
        ///     Called when a right-click is performed in TIA
        ///     and a mouse-over is performed on the name of the Add-In.
        /// </summary>
        /// <typeparam name="addInRootSubmenu">
        ///     The Add-In will be displayed in
        ///     the Context Menu of TIA Portal.
        /// </typeparam>
        /// <example>
        ///     ActionItems like Buttons/Checkboxes/Radiobuttons
        ///     are possible. In this example, only Buttons will be created
        ///     which will start the Add-In program code.
        /// </example>
        protected override void BuildContextMenuItems(ContextMenuAddInRoot
            addInRootSubmenu)
        {
            /* Method addInRootSubmenu.Items.AddActionItem
            * Will Create a Pushbutton with the text 'Start Add-In Code'
            * 1st input parameter of AddActionItem is the text of the
            * button
            * 2nd input parameter of AddActionItem is the clickDelegate,
            * which will be executed in case the button 'Start
            * Add-In Code' will be clicked/pressed.
            * 3rd input parameter of AddActionItem is the
            * updateStatusDelegate, which will be executed in
            * case there is a mouseover the button 'Start
            * Add-In Code'.
            * in <placeholder> the type of AddActionItem will be
            * specified, because AddActionItem is generic
            * AddActionItem<DeviceItem> will create a button that will be
            * displayed if a rightclick on a DeviceItem will be
            * performed in TIA Portal
            * AddActionItem<Project> will create a button that will be
            * displayed if a rightclick on the project name
            * will be performed in TIA Portal
            */
            //addInRootSubmenu.Items.AddActionItem<DeviceItem>(
            //("Start Add-In"), OnDoSomething, OnCanSomething);

            //addInRootSubmenu.Items.AddActionItem<Project>(
            //"Not Available here", OnClickProject,
            //OnStatusUpdateProject);
            addInRootSubmenu.Items.AddActionItem<DeviceItem>("Start Simulation", OnStartSimulation,
                OnStartSimulationUpdateStatus);
            addInRootSubmenu.Items.AddActionItem<DeviceItem>("Stop Simulation", OnStopSimulation,
                OnStopSimulationUpdateStatus);
        }

        /// <summary>
        /// Main function to start the addin. The call starts the .exe process.
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void OnStartSimulation(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            DemandProcessStartPermission();
         

            //var folderPath = GetCommonApplicationDataFolderPath();
            //Directory.CreateDirectory(folderPath);
            //var exePath = Path.Combine(folderPath, "PLCSIMAdvStarter.exe");

            // Form "PLCSIMAdv-Starter.exe" from Resources to the specified folderPath if not created yet
            //if (!Directory.GetFiles(folderPath).Any())
            //{
            //    // Extract Server Executable
            //    var exe = Resource.PLCSIMAdvStarter;
            //    File.WriteAllBytes(exePath, exe);
            //}

            // Fetch plcName and processID of the running TIA Portal
            var deviceItem = GetDeviceItem(menuSelectionProvider);
            var plcName = deviceItem.Name;
            var tiaProcessID = _tiaportal.GetCurrentProcess().Id.ToString();
            string folderPath = GetExeAndStateFileLocation();


            //PersistStateFile(folderPath, c_StartState, plcName);

            // Start "PLCSIMAdv-Starter.exe" process if no StateFile created yet
            //if (File.Exists(Path.Combine(folderPath, "StateFile.txt")) && File.ReadLines(Path.Combine(folderPath, "StateFile.txt")).Any(line => line.Contains(c_StartState)))
            //{
            CliHandling.RunExecutable(plcName, tiaProcessID, folderPath);
                //Process.Start(exePath, plcName + $" {tiaProcessID}");
                //Process.Start(exePath);
            //}

            

            PersistStateFile(folderPath, c_StartState, plcName);
        }


        /// <summary>
        /// Gets the location of the .exe file and the state file
        /// </summary>
        /// <returns></returns>
        private string GetExeAndStateFileLocation()
        {
            var asm = Assembly.GetExecutingAssembly();
            var companyName = (asm.GetCustomAttribute(typeof(AssemblyCompanyAttribute)) as AssemblyCompanyAttribute)
                .Company;
            var applicationName = (asm.GetCustomAttribute(typeof(AssemblyTitleAttribute)) as AssemblyTitleAttribute)
                .Title;
            var baseFolder =
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    companyName, "Automation", applicationName);
            return baseFolder;
        }

        /// <summary>
        /// Stops the simulation process
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        private void OnStopSimulation(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
            var folderPath = GetExeAndStateFileLocation();
            // Fetch plcName
            var deviceItem = GetDeviceItem(menuSelectionProvider);
            var plcName = deviceItem.Name;

            PersistStateFile(folderPath, c_StopState, plcName);
        }

        /// <summary>
        /// Function creates the StateFile which contains the actual state of the simulation process. 
        /// </summary>
        /// <param name="folderPath"></param>
        /// <param name="state"></param>
        /// <param name="plcName"></param>
        private void PersistStateFile(string folderPath, string state, string plcName)
        {
            var filePath = Path.Combine(folderPath, "StateFile.txt");

            if (false == File.Exists(filePath))
            {
                var stream = File.Open(filePath, FileMode.OpenOrCreate);
                stream.Close();
            }

            using (var outStream = new StreamWriter(File.Open(filePath, FileMode.Truncate,
                       FileAccess.Write,
                       FileShare.ReadWrite)))
            {
                outStream.Write("");
                outStream.Write($"{state},{plcName}");
                outStream.Close();
            }
        }


        /// <summary>
        /// Create Directory 'PLCSIMAdv' under CommonApplicationData directory
        /// </summary>
        /// <returns></returns>
        private string GetCommonApplicationDataFolderPath()
        {
            var path = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);
            path = Path.Combine(path, "PLCSIMAdvAddIn");
            return path;
        }

        /// <summary>
        /// Delegate if the start simulation is visible at the project tree (TIA Portal)
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        private MenuStatus OnStartSimulationUpdateStatus(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
            var deviceItem = GetDeviceItem(menuSelectionProvider);
            var onlineProvider = deviceItem.GetService<OnlineProvider>();

            return onlineProvider.State == OnlineState.Offline ? MenuStatus.Enabled : MenuStatus.Disabled;
        }

        /// <summary>
        /// Delegate if the stop simulation is visible at the project tree (TIA Portal)
        /// </summary>
        /// <param name="menuSelectionProvider"></param>
        /// <returns></returns>
        private MenuStatus OnStopSimulationUpdateStatus(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
            var deviceItem = GetDeviceItem(menuSelectionProvider);
            var onlineProvider = deviceItem.GetService<OnlineProvider>();

            return onlineProvider.State == OnlineState.Online ? MenuStatus.Enabled : MenuStatus.Disabled;
        }

        private static DeviceItem GetDeviceItem(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
            return menuSelectionProvider.GetSelection<DeviceItem>().First();
        }

        private void DemandProcessStartPermission()
        {
            new ProcessStartPermission(PermissionState.Unrestricted).Demand();
        }
    }
}