using System;
using System.Diagnostics;
using System.IO;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;

namespace gitSelect2
{
    internal class Program
    {
        private static readonly string logFilePath = @"C:\git3\Progs\log.txt";
        private static readonly string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ubiAccess.json");

        static void Main(string[] args)
        {
            try
            {
                LogMessage("Empezando el proceso...");

                // Define the two paths where the version.txt files are located
                string pathLocal = @"C:\git3\Progs\version.txt";
                string pathServer = @"S:\git3\Progs\version.txt";
                //string pathServer = @"C:\git3\Progs\version2.txt";

                string[] filesToCopy =
                {
                    "gitalm.ade",
                    "gitcom.ade",
                    "gitcon.ade",
                    "gitmaes.ade",
                    "gitmenu.ade",
                    "gitped.ade",
                    "gitpro.ade",
                    "gitven.ade"
                };

                // Read the version numbers from both files
                string versionLocal = File.ReadAllText(pathLocal).Trim();
                string versionServer = File.ReadAllText(pathServer).Trim();

                // Compare the version numbers
                if (versionLocal == versionServer)
                {
                    LogMessage("Los números de versión son iguales.");
                }
                else
                {
                    LogMessage("Los números de versión son diferentes.");

                    // Show a message box using PowerShell
                    bool userWantsToUpdate = ShowConfirmationDialog();

                    if (userWantsToUpdate)
                    {
                        // Copy the version from the server file to the local file
                        File.WriteAllText(pathLocal, versionServer);
                        LogMessage($"La versión de {pathServer} se ha copiado a {pathLocal}.");

                        // Copy the specified files from the server directory to the local directory
                        string localDirectory = Path.GetDirectoryName(pathLocal);
                        string serverDirectory = Path.GetDirectoryName(pathServer);

                        foreach (string fileName in filesToCopy)
                        {
                            string serverFilePath = Path.Combine(serverDirectory, fileName);
                            string localFilePath = Path.Combine(localDirectory, fileName);

                            if (File.Exists(serverFilePath))
                            {
                                File.Copy(serverFilePath, localFilePath, true);
                                LogMessage($"El archivo {fileName} se ha copiado de {serverDirectory} a {localDirectory}.");
                            }
                            else
                            {
                                LogMessage($"El archivo {fileName} no existe en el servidor.");
                            }
                        }
                    }
                    else
                    {
                        LogMessage("Actualización cancelada por el usuario.");
                    }
                }

                // Execute MSACCESS.EXE
                ExecuteMsAccess();

            }
            catch (Exception ex)
            {
                // Handle any errors that occur during file reading or writing
                LogMessage("Ocurrió un error: " + ex.Message);
            }

            Console.WriteLine("Proceso terminado. Presiona cualquier tecla para salir.");
            Console.ReadKey();
        }

        private static bool ShowConfirmationDialog()
        {
            string script = @"
                Add-Type -AssemblyName System.Windows.Forms
                $result = [System.Windows.Forms.MessageBox]::Show('Hay una actualización disponible. ¿Desea actualizar los archivos locales con los del servidor?', 'Confirmación de actualización', 'YesNo', 'Question')
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { exit 0 } else { exit 1 }
            ";

            ProcessStartInfo startInfo = new ProcessStartInfo()
            {
                FileName = "powershell",
                Arguments = $"-NoProfile -Command \"{script}\"",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(startInfo))
            {
                process.WaitForExit();
                return process.ExitCode == 0;
            }
        }

        private static void ExecuteMsAccess()
        {
            try
            {
                string msAccessPath = GetMsAccessPath();

                if (!string.IsNullOrEmpty(msAccessPath) && File.Exists(msAccessPath))
                {
                    string arguments = $"/runtime \"c:\\git3\\Progs\\GITMENU.ade\"";
                    Process.Start(msAccessPath, arguments);
                    LogMessage($"MSACCESS.EXE ejecutado desde: {msAccessPath} con argumentos: {arguments}");
                }
                else
                {
                    LogMessage("No se encontró MSACCESS.EXE en el sistema.");
                }
            }
            catch (Exception ex)
            {
                LogMessage("Ocurrió un error al intentar ejecutar MSACCESS.EXE: " + ex.Message);
            }
        }

        private static string GetMsAccessPath()
        {
            string msAccessPath = string.Empty;

            if (File.Exists(jsonFilePath))
            {
                string jsonContent = File.ReadAllText(jsonFilePath);
                JObject jsonObj = JObject.Parse(jsonContent);

                msAccessPath = jsonObj["ubicacionAccess"]?.ToString();

                if (!string.IsNullOrEmpty(msAccessPath) && File.Exists(msAccessPath))
                {
                    return msAccessPath;
                }
            }

            // Search for MSACCESS.EXE if not found in JSON or if the path is invalid
            msAccessPath = SearchForMsAccess();

            // Update the JSON file with the found path
            if (!string.IsNullOrEmpty(msAccessPath))
            {
                JObject jsonObj = new JObject
                {
                    ["ubicacionAccess"] = msAccessPath
                };
                File.WriteAllText(jsonFilePath, jsonObj.ToString());
                LogMessage($"Ubicación de MSACCESS.EXE guardada en {jsonFilePath}");
            }

            return msAccessPath;
        }

        private static string SearchForMsAccess()
        {
            string[] potentialPaths =
            {
                // Office 2016, Office 2019, Office 2021 (32-bit y 64-bit)
                @"C:\Program Files\Microsoft Office\root\Office16",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16",
                @"C:\Program Files\Microsoft Office\root\Office19",
                @"C:\Program Files (x86)\Microsoft Office\root\Office19",
                @"C:\Program Files\Microsoft Office\root\Office21",
                @"C:\Program Files (x86)\Microsoft Office\root\Office21",
    
                // Office 2013 (32-bit y 64-bit)
                @"C:\Program Files\Microsoft Office\Office15",
                @"C:\Program Files (x86)\Microsoft Office\Office15",

                // Office 2010 (32-bit y 64-bit)
                @"C:\Program Files\Microsoft Office\Office14",
                @"C:\Program Files (x86)\Microsoft Office\Office14",

                // Office 2007 (32-bit y 64-bit)
                @"C:\Program Files\Microsoft Office\Office12",
                @"C:\Program Files (x86)\Microsoft Office\Office12",

                // Office 2003 (32-bit)
                @"C:\Program Files\Microsoft Office\Office11",
                @"C:\Program Files (x86)\Microsoft Office\Office11",

                // Office 2000 (32-bit)
                @"C:\Program Files\Microsoft Office\Office",
                @"C:\Program Files (x86)\Microsoft Office\Office",

                // Office 365 (32-bit y 64-bit)
                @"C:\Program Files\Microsoft Office\root\Office16",
                @"C:\Program Files (x86)\Microsoft Office\root\Office16",
                @"C:\Program Files\Microsoft Office\root\Office19",
                @"C:\Program Files (x86)\Microsoft Office\root\Office19",
                @"C:\Program Files\Microsoft Office\root\Office21",
                @"C:\Program Files (x86)\Microsoft Office\root\Office21",

                // Versiones de Office instaladas en ubicaciones personalizadas
                @"D:\Program Files\Microsoft Office\root\Office16",
                @"D:\Program Files (x86)\Microsoft Office\root\Office16",
                @"D:\Program Files\Microsoft Office\Office15",
                @"D:\Program Files (x86)\Microsoft Office\Office15",
                @"D:\Program Files\Microsoft Office\Office14",
                @"D:\Program Files (x86)\Microsoft Office\Office14",
                @"D:\Program Files\Microsoft Office\Office12",
                @"D:\Program Files (x86)\Microsoft Office\Office12",
                @"D:\Program Files\Microsoft Office\Office11",
                @"D:\Program Files (x86)\Microsoft Office\Office11"
            };

            foreach (var path in potentialPaths)
            {
                Console.WriteLine($"Verificando ruta: {path}");

                if (Directory.Exists(path))
                {
                    var exePath = Directory.GetFiles(path, "MSACCESS.EXE", SearchOption.AllDirectories);
                    if (exePath.Length > 0)
                    {
                        Console.WriteLine($"Encontrado: {exePath[0]}");
                        return exePath[0];
                    }
                    else
                    {
                        Console.WriteLine($"MSACCESS.EXE no encontrado en: {path}");
                    }
                }
                else
                {
                    Console.WriteLine($"Ruta no existe: {path}");
                }
            }

            Console.WriteLine("MSACCESS.EXE no encontrado en ninguna de las rutas especificadas.");
            return string.Empty;
        }

        private static void LogMessage(string message)
        {
            try
            {
                // Ensure the directory exists
                string logDirectory = Path.GetDirectoryName(logFilePath);
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                // Append the message to the log file
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                }
                Console.WriteLine(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error al escribir en el archivo de log: " + ex.Message);
            }
        }
    }
}