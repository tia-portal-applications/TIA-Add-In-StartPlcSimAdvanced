using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using Siemens.Engineering.AddIn.Utilities;

namespace PLCSIMAdvAddIn
{
    internal class CliHandling
    {
        private const string ExecutablePath = "Delivery/PLCSIMAdvStarter.exe";

        /// <summary>
        /// Main Function to get the .exe file
        /// </summary>
        /// <param name="plcName"></param>
        /// <param name="processID"></param>
        /// <param name="folderPath"></param>
        public static void RunExecutable(string plcName, string processID, string folderPath)
        {
            var cliPath = GetOrExtractExecutable();

            var startInfo = new ProcessStartInfo
            {
                FileName = cliPath,
                Arguments = plcName + $" {processID}" + $" {folderPath}",
                CreateNoWindow = true,
                UseShellExecute = false
            };

            var cliProcess = new Process
            {
                EnableRaisingEvents = true,
                StartInfo = startInfo
            };
            cliProcess.Start();
        }

        /// <summary>
        /// Gets or extracts the executable
        /// </summary>
        /// <returns></returns>
        private static string GetOrExtractExecutable()
        {
            var asm = Assembly.GetExecutingAssembly();
            var companyName = (asm.GetCustomAttribute(typeof(AssemblyCompanyAttribute)) as AssemblyCompanyAttribute)
                .Company;
            var applicationName = (asm.GetCustomAttribute(typeof(AssemblyTitleAttribute)) as AssemblyTitleAttribute)
                .Title;
            var baseFolder =
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    companyName, "Automation", applicationName);

            if (!Directory.Exists(baseFolder)) Directory.CreateDirectory(baseFolder);

            var targetPath = Path.Combine(baseFolder, ExecutablePath);

            if (File.Exists(targetPath))
            {
                if (CheckExistingFileIntegrity(targetPath))
                    return targetPath;
                Directory.Delete(Path.GetDirectoryName(targetPath), true);
            }

            var tmpZipPath = Path.Combine(baseFolder, "AddInContainedExe.zip");
            var assembly = Assembly.GetExecutingAssembly();
            using (var resource = assembly.GetManifestResourceStream("PLCSIMAdv.ExecutableDelivery.zip"))
            {
                using (var file = new FileStream(tmpZipPath, FileMode.Create, FileAccess.ReadWrite))
                {
                    resource.CopyTo(file);
                }
            }

            ZipFile.ExtractToDirectory(tmpZipPath, baseFolder);
            File.Delete(tmpZipPath);

            return targetPath;
        }

        /// <summary>
        /// Checks if there is a new executable file
        /// </summary>
        /// <param name="existingFilePath"></param>
        /// <returns></returns>
        private static bool CheckExistingFileIntegrity(string existingFilePath)
        {
            var shaExistingFile = GetSha256(existingFilePath);

            var assembly = Assembly.GetExecutingAssembly();
            var shaCompare =
                new StreamReader(assembly.GetManifestResourceStream("PLCSIMAdv.ExecutableChecksum.txt"))
                    .ReadToEnd();
            shaCompare = shaCompare.Trim();

            return shaExistingFile.Equals(shaCompare, StringComparison.CurrentCultureIgnoreCase);
        }

        /// <summary>
        /// Creates a SHA256 key so compare it 
        /// </summary>
        /// <param name="inputPath"></param>
        /// <returns></returns>
        private static string GetSha256(string inputPath)
        {
            using (var stream = File.OpenRead(inputPath))
            {
                var sha = new SHA256Managed();
                var checksum = sha.ComputeHash(stream);
                return BitConverter.ToString(checksum).Replace("-", string.Empty);
            }
        }
    }
}