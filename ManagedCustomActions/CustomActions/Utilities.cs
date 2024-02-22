using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.IO;
using System.Security.Principal;

namespace CustomActions
{
   public partial class CustomActions
   {
      // InstallWizard
      internal static Seagull.InstallWizard.InstallWizard wizard = new Seagull.InstallWizard.InstallWizard();

      private const int SW_HIDE = 0;
      private const int SW_SHOWNORMAL = 1;
      private const int SW_SHOW = 5;
      private const int SW_MINIMIZE = 6;
      private const int SW_RESTORE = 9;
      private const int OS_ANYSERVER = 29;

      internal const string defaultSQLInstanceName = "BarTender";

      // DLL Imports
      [DllImport("shlwapi.dll", SetLastError = true, EntryPoint = "#437")]
      private static extern bool IsOS(int os);

      [DllImport("user32.dll")]
      [return: MarshalAs(UnmanagedType.Bool)]
      private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, ref SearchData data);

      [DllImport("user32.dll", SetLastError = true)]
      static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

      [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
      public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

      [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
      public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

      [DllImport("User32")]
      public static extern int ShowWindow(IntPtr hwnd, int nCmdShow);

      // List of all known installer features
      public enum InstallFeatures
      {
         AdministrationConsole,
         BPP,
         BarTender,
         Common,
         DataBuilder,
         HistoryExplorer,
         IntegrationBuilder,
         Languages,
         Librarian,
         PrintStation,
         PrinterMaestro,
         ProcessBuilder,
         ReprintConsole,
         SDK,
         SystemDatabase,
         PrintRouter,
         MicrosoftNETCore,
         MicrosoftNETCore_Repair
      }

      // Mapped sets of InstallFeatures
      internal static Dictionary<string, List<InstallFeatures>> MappedInstallFeatures = new Dictionary<string, List<InstallFeatures>>(StringComparer.OrdinalIgnoreCase)
      {
         { "BarTender", new List<InstallFeatures>() {
               InstallFeatures.AdministrationConsole,
               InstallFeatures.BarTender,
               InstallFeatures.Common,
               InstallFeatures.DataBuilder,
               InstallFeatures.HistoryExplorer,
               InstallFeatures.IntegrationBuilder,
               InstallFeatures.Languages,
               InstallFeatures.Librarian,
               InstallFeatures.PrintStation,
               InstallFeatures.PrinterMaestro,
               InstallFeatures.ProcessBuilder,
               InstallFeatures.ReprintConsole,
               InstallFeatures.SDK,
               InstallFeatures.SystemDatabase
         }},
         { "PrintPortal", new List<InstallFeatures>() {
               InstallFeatures.AdministrationConsole,
               InstallFeatures.BPP,
               InstallFeatures.BarTender,
               InstallFeatures.Common,
               InstallFeatures.DataBuilder,
               InstallFeatures.HistoryExplorer,
               InstallFeatures.IntegrationBuilder,
               InstallFeatures.Languages,
               InstallFeatures.Librarian,
               InstallFeatures.PrintStation,
               InstallFeatures.PrinterMaestro,
               InstallFeatures.ProcessBuilder,
               InstallFeatures.ReprintConsole,
               InstallFeatures.SDK,
               InstallFeatures.SystemDatabase,
               InstallFeatures.PrintRouter,
               InstallFeatures.MicrosoftNETCore,
               InstallFeatures.MicrosoftNETCore_Repair
         }},
         { "LicensingService", new List<InstallFeatures>() {
               InstallFeatures.AdministrationConsole,
               InstallFeatures.Common,
               InstallFeatures.Languages,
               InstallFeatures.SystemDatabase
         }}
      };

      // Method to convert features being added and removed (CSV formatted) into a list of InstallFeatures
      internal static List<InstallFeatures> GetFeatures(string addFeatures, string removeFeatures)
      {
         // Parse addFeatures
         List<InstallFeatures> addList = new List<InstallFeatures>();
         string[] addFeaturesArray = addFeatures.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
         foreach (string feature in addFeaturesArray)
         {
            // When only one item and it's in the MappedInstallFeatures, use it instead
            if (addFeaturesArray.Length == 1 && MappedInstallFeatures.ContainsKey(feature))
            {
               addList.AddRange(MappedInstallFeatures[feature]);
               continue;
            }

            // Add to list
            InstallFeatures feat;
            if (Enum.TryParse<InstallFeatures>(feature, true, out feat))
               addList.Add(feat);
         }

         // Parse removeFeatures
         List<InstallFeatures> removeList = new List<InstallFeatures>();
         string[] removeFeaturesArray = removeFeatures.Split(",".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
         foreach (string feature in removeFeaturesArray)
         {
            // Add to list
            InstallFeatures feat;
            if (Enum.TryParse<InstallFeatures>(feature, true, out feat))
               removeList.Add(feat);
         }

         // Subtract the removeList from the addList and return
         addList.RemoveAll(r => removeList.Contains(r));
         return addList;
      }

      // SearchData for EnumWindows
      public class SearchData
      {
         public string Wndclass;
         public string Title;
         public IntPtr hWnd;
      }

      // Wraps the window handle for IWin32Window interface
      public class WindowWrapper : System.Windows.Forms.IWin32Window
      {
         public WindowWrapper(IntPtr handle)
         {
            _hwnd = handle;
         }

         public IntPtr Handle
         {
            get { return _hwnd; }
         }

         private IntPtr _hwnd;
      }

      // Find and fetch main window handle by process name
      public static IntPtr GetWindowHandleByProcessName(string ProcessName)
      {
         IntPtr pFoundWindow = IntPtr.Zero;
         try
         {
            foreach (Process p in Process.GetProcessesByName(ProcessName))
               if (p.MainWindowHandle != IntPtr.Zero)
                  pFoundWindow = p.MainWindowHandle;
         }
         catch { }
         return pFoundWindow;
      }

      private delegate bool EnumWindowsProc(IntPtr hWnd, ref SearchData data);

      public static bool EnumProc(IntPtr hWnd, ref SearchData data)
      {
         // Check classname and title
         // This is different from FindWindow() in that the code below allows partial matches
         StringBuilder sb = new StringBuilder(1024);
         GetClassName(hWnd, sb, sb.Capacity);
         if (sb.ToString().StartsWith(data.Wndclass))
         {
            sb = new StringBuilder(1024);
            GetWindowText(hWnd, sb, sb.Capacity);
            if (sb.ToString().StartsWith(data.Title))
            {
               data.hWnd = hWnd;
               return false;    // Found the wnd, halt enumeration
            }
         }
         return true;
      }

      public static IntPtr SearchForWindow(string wndclass, string title)
      {
         SearchData sd = new SearchData { Wndclass = wndclass, Title = title };
         EnumWindows(new EnumWindowsProc(EnumProc), ref sd);
         return sd.hWnd;
      }

      // Check the version of IIS on the machine
      private static string IISVersion()
      {
         string iisVersion = string.Empty;
         if (isIIS())
         {
            // Find IIS version
            RegistryKey IISVersionKey = Registry.LocalMachine;
            try
            {
               IISVersionKey = IISVersionKey.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\W3SVC\Parameters", RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.QueryValues);
               iisVersion = IISVersionKey.GetValue("MajorVersion").ToString();
            }
            catch { }
         }
         return iisVersion;
      }

      // Check to see if IIS is installed on the machine.
      private static bool isIIS()
      {
         // Find IIS location in the registry
         string IISkey = string.Empty;
         RegistryKey IISRegKey = Registry.LocalMachine;

         try
         {
            IISRegKey = IISRegKey.OpenSubKey(@"Software\Microsoft\InetStp", RegistryKeyPermissionCheck.ReadSubTree, System.Security.AccessControl.RegistryRights.QueryValues);
            IISkey = IISRegKey.GetValue("PathWWWRoot").ToString();
         }
         catch { }

         if (!string.IsNullOrEmpty(IISkey))
         {
            return true;
         }
         else
         {
            return false;
         }
      }

      // Start a service, returns true if and only if the service was started
      private static bool StartService(string serviceName)
      {
         return StartServices( new List<string>(){ serviceName });
      }

      // Starts requested services and returns true if and only if all of the services are started up
      private static bool StartServices(List<string> serviceNames)
      {
         bool retval = true;
         foreach( string serviceName in serviceNames )
         {
            try
            {
               // Check if service exists
               ServiceController service = null;
               try
               {
                  service = ServiceController.GetServices().First( r => r.DisplayName == serviceName);
                  if (service == null)
                     throw new InvalidOperationException("Unable to find service by DisplayName");
               }
               catch (InvalidOperationException)
               {
                  // Service doesn't exist
                  retval = false;
                  continue;
               }

               // Check status and start if needed
               if (service.Status == ServiceControllerStatus.Stopped)
               {
                  if (new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
                  {
                     service.Start();
                  }
                  else
                  {
                     System.Diagnostics.Process process = new System.Diagnostics.Process();
                     process.StartInfo.UseShellExecute = true;
                     process.StartInfo.Verb = "runas";
                     process.StartInfo.RedirectStandardOutput = false;
                     process.StartInfo.RedirectStandardError = false;
                     process.StartInfo.CreateNoWindow = true;

                     process.StartInfo.FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "sc.exe");
                     process.StartInfo.WorkingDirectory = Environment.GetFolderPath(Environment.SpecialFolder.System);
                     process.StartInfo.Arguments = "start \"" + serviceName + "\"";
                     process.Start();
                     process.WaitForExit(15000);
                  }
               }
               else if (service.Status == ServiceControllerStatus.StartPending)
               {
                  // Wait up to X sec for service to finish starting
                  try
                  {
                     TimeSpan timeout = TimeSpan.FromSeconds(30);
                     service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                  }
                  catch { }
               }
               else
               {
                  // Service is not in a stopped or startpending state, therefore do nothing to start it up
               }

               // Check that service is now running
               service.Refresh();
               if (service.Status != ServiceControllerStatus.Running)
                  retval = false;
            }
            catch
            {
               // handle exceptions
               retval = false;
               throw;
            }
         }

         // Return overall result after having tried to start all services
         return retval;
      }


      /// <summary>
      /// Helper method to load session data into wizard and santize session data
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      internal static void PrepareWizard(Session session)
      {
         // InstallPath
         if (session["INSTALLDIR"].Trim() != string.Empty)
         {
            wizard.InstallPath = session["INSTALLDIR"].Trim();
         }

         // Feature
         if (session["FEATURE"].Trim() != string.Empty)
         {
            Seagull.InstallWizard.InstallTypeSelection setInstallType;
            if (Enum.TryParse<Seagull.InstallWizard.InstallTypeSelection>(session["FEATURE"].ToString(), true, out setInstallType))
            {
               wizard.InstallType = setInstallType;
            }
            else
            {
               session.Log("Warning: an error occured parsing the FEATURE property, wizard will use default install type instead");
            }
         }

         // Product Language
         wizard.InstallLanguage = session["ProductLanguage"];

         // Our version info
         wizard.MajorVersion = session["BT_PRODUCT_VERSION"];
         wizard.FullVersion = session["FullProductVersion"];

         // SQL Instance + SystemDB states
         CustomActions.GetSQLInstanceState(session);
         wizard.SQLInstalled = (session["SQLInstalled"] == "true");
         CustomActions.GetSystemDBState(session);
         wizard.SystemDbIsSetup = (session["SystemDbIsSetup"] == "true");

         // Existing version info
         wizard.InstalledVersion = session["BT_INSTALLED_VERSION_STRING"];
         string currentInstalledVersion = session["BT_INSTALLED_VERSION"].Trim();
         if (!string.IsNullOrEmpty(currentInstalledVersion))
         {
            Version installerVersion = new Version();
            Version currentVersion = new Version();

            session.Log(string.Format("Currently installed version: {0}", currentInstalledVersion));
            session.Log(string.Format("Installer's version: {0}", wizard.FullVersion));
            try
            {
               installerVersion = new Version(wizard.FullVersion);
               currentVersion = new Version(currentInstalledVersion);
            }
            catch (Exception ex)
            {
               session.Log("Exception has occurred:");
               session.Log(ex.Message);
            }

            if (currentVersion != new Version() && installerVersion != new Version())
            {
               if (installerVersion == currentVersion)
               {
                  // Versions are the same
                  session.Log(string.Format("Installer versions match"));
               }
               else if (installerVersion > currentVersion)
               {
                  if ((session["x86_or_x64"] == "x86" && Environment.Is64BitOperatingSystem) || (session["x86_or_x64"] == "x64" && !Environment.Is64BitOperatingSystem))
                  {
                     // Installer is a cross platform upgrade, set AI_UPGRADE to trigger UNINSTALL_PREVINSTALL_XPLATFORM_UPGRADE
                     session.Log(string.Format("This installer is a cross-platform upgrade"));
                     session["AI_UPGRADE"] = "Yes";
                     wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.OlderVersion;
                  }
                  else
                  {
                     session.Log(string.Format("This installer is an upgrade"));
                     session["AI_UPGRADE"] = "Yes";
                     session["MIGRATE"] = session["OLDPRODUCTS"];
                     // Installer is newer, so tell wizard that the previous version is 'older'
                     wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.OlderVersion;
                  }
               }
               else if (installerVersion < currentVersion)
               {
                  // Installer is older, so tell wizard that the previous version is 'newer'
                  session.Log(string.Format("This installer is a downgrade"));
                  session["AI_DOWNGRADE"] = "Yes";
                  session["REINSTALLMODE"] = "amus";
                  wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.NewerVersion;
               }
            }
            else
            {
               session.Log("Unable to parse either the current installation's or installer's version numbers");
            }
         }
      }

      // Checks if Windows Update has a pending reboot
      internal static void GetWindowsUpdateRebootRequired(Session session)
      {
         try
         {
            RegistryKey rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            using (RegistryKey key = rk.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"))
            {
               foreach (string valueName in key.GetValueNames())
               {
                  if (valueName.Trim() != "")
                     session["WindowsUpdateRebootRequired"] = "true";
               }
            }
         }
         catch (Exception e)
         {
            session.Log("Error while handling registry request: " + e.Message);
         }
         finally
         {
            if (session["WindowsUpdateRebootRequired"] != "true")
               session["WindowsUpdateRebootRequired"] = "false";
         }
      }

      // Checks if SQL instance exists on the local system
      internal static void GetSQLInstanceState(Session session)
      {
         try
         {
            RegistryKey rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            using (RegistryKey key = rk.OpenSubKey( @"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL" ))
            {
               if (key != null && key.GetValue(defaultSQLInstanceName) != null)
                  session["SQLInstalled"] = "true";
            }
            using (RegistryKey key = rk.OpenSubKey( @"SOFTWARE\WOW6432Node\Microsoft\Microsoft SQL Server\Instance Names\SQL" ))
            {
               if (key != null && key.GetValue(defaultSQLInstanceName) != null)
                  session["SQLInstalled"] = "true";
            }
         }
         catch (Exception e)
         {
            session.Log("Error while handling registry request: " + e.Message);
         }
         finally
         {
            if (session["SQLInstalled"] != "true")
               session["SQLInstalled"] = "false";
         }
      }

      // Checks if SQL instance exists on the local system
      internal static void GetSystemDBState(Session session)
      {
         try
         {
            Assembly btSystemClient = null;
            Type serviceClientHelper = null;
            String assemblyFilePath = "";
            String assemblyInfo = "";

            // Search likely install paths containing a BtSystem.Client.dll
            if (File.Exists(session["APPDIR"] + "\\BtSystem.Client.dll"))
               assemblyFilePath = session["APPDIR"] + "\\BtSystem.Client.dll";
            if (File.Exists(session["BT_INSTALLED_INSTALLLOCATION"] + "\\BtSystem.Client.dll"))
               assemblyFilePath = session["BT_INSTALLED_INSTALLLOCATION"] + "\\BtSystem.Client.dll";

            // Try using install path's BtSystem.Client
            if (assemblyFilePath != "")
            {
               try
               {
                  // LoadFile Assembly
                  btSystemClient = Assembly.LoadFile(assemblyFilePath);
                  serviceClientHelper = btSystemClient.GetType("Seagull.BtSystem.Client.ServiceClientHelper");
                  session.Log("BtSystem.Client loaded from \"" + assemblyFilePath + "\"");
               }
               catch (Exception e)
               {
                  session.Log("Exception: unable to load BtSystem.Client from \"" + assemblyFilePath + "\" - " + e.Message);
               }
            }
            
            // If install path didn't work, try finding BtSystem.Client in the GAC
            if (serviceClientHelper == null)
            {
               try
               {
                  // Search GAC for BtSystem.Client
                  assemblyInfo = GacApi.QueryAssemblyInfo("BtSystem.Client");   // Returns full path to assembly
                  assemblyInfo = Assembly.ReflectionOnlyLoadFrom(assemblyInfo).FullName;  // Strong name

                  // Load Assembly
                  btSystemClient = Assembly.Load(assemblyInfo);
                  serviceClientHelper = btSystemClient.GetType("Seagull.BtSystem.Client.ServiceClientHelper");
                  session.Log("BtSystem.Client loaded from \"" + assemblyInfo + "\"");
               }
               catch (Exception e)
               {
                  session.Log("Exception: unable to load BtSystem.Client from GAC \"" + assemblyInfo + "\" - " + e.Message);
               }
            }

            // Return now if no ServiceClientHelper was available
            if (serviceClientHelper == null)
               return;

            // Create an instance of the ServiceClientHelper object, and query for the values we're interested in
            object schInstance = Activator.CreateInstance(serviceClientHelper, new object[] { });
            session["SystemDbIsSetup"] = ((bool)serviceClientHelper.GetProperty("SystemDbIsSetup").GetValue(schInstance, null)) ? "true" : "false";
            session["SystemDbIsDbUpgradeRequired"] = ((bool)serviceClientHelper.GetMethod("IsDbUpgradeRequired").Invoke(schInstance, new object[] { "" })) ? "true" : "false";
         }
         catch (System.IO.FileNotFoundException)
         {
            session.Log("Failed to find a BtSystem.Client");
         }
         catch (Exception e)
         {
            session.Log("Unhandled Exception: " + e.Message);
         }
      }

      /// <summary>
      /// Run a command line, and capture both of its standard out and standard error to a string.
      /// </summary>
      /// <param name="command">Command to execute.</param>
      /// <param name="parameters">Arguments for the command.</param>
      /// <returns>String containing the captured outputs.</returns>
      public static string RunCommandAndCaptureAllOutputs(string command, string parameters, int timeout = System.Threading.Timeout.Infinite)
      {
         // Start the child process.
         Process p = new Process();

         // Setup an instance of the asynchronous stream reader
         AsyncStreamReader async = new AsyncStreamReader();

         // Redirect the output and error streams of the child process.
         p.StartInfo.FileName = command;
         p.StartInfo.Arguments = parameters;
         p.StartInfo.UseShellExecute = false;
         p.StartInfo.RedirectStandardOutput = true;
         p.OutputDataReceived += new DataReceivedEventHandler(async.StdOutHandler);
         p.StartInfo.RedirectStandardError = true;
         p.ErrorDataReceived += new DataReceivedEventHandler(async.StdErrHandler);
         p.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
         p.StartInfo.CreateNoWindow = true;

         // Start process
         p.Start();

         // Begin reading streams
         p.BeginOutputReadLine();
         p.BeginErrorReadLine();

         // Block until process exits
         p.WaitForExit(timeout);

         // Combine streams into one return value
         string retval = async.stdOutput.ToString();
         if (async.stdError.Length > 0)
         {
            retval += Environment.NewLine + "STANDARD ERROR:" + Environment.NewLine;
            retval += async.stdError.ToString();
         }

         p.Close();
         return retval;
      }

      private class AsyncStreamReader
      {
         // strings used for async process handler
         public StringBuilder stdOutput = new StringBuilder("");
         public StringBuilder stdError = new StringBuilder("");

         public void StdOutHandler(object sendingProcess, DataReceivedEventArgs lineOut)
         {
            // Add to output string
            if (!String.IsNullOrEmpty(lineOut.Data))
            {
               stdOutput.AppendLine(lineOut.Data);
            }
         }

         public void StdErrHandler(object sendingProcess, DataReceivedEventArgs lineOut)
         {
            // Add to error string
            if (!String.IsNullOrEmpty(lineOut.Data))
            {
               stdError.AppendLine(lineOut.Data);
            }
         }
      }
   }

   // Class for querying Global Assembly Cache
   internal class GacApi
   {
      [DllImport("fusion.dll")]
      internal static extern IntPtr CreateAssemblyCache(out IAssemblyCache ppAsmCache, int reserved);

      // IAssemblyCache interface (partial, with non-used vtable entries)
      [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("e707dcde-d1cd-11d2-bab9-00c04f8eceae")]
      internal interface IAssemblyCache
      {
         int Dummy1();
         [PreserveSig()]
         IntPtr QueryAssemblyInfo(
            int flags,
            [MarshalAs(UnmanagedType.LPWStr)] 
         String assemblyName,
            ref AssemblyInfo assemblyInfo);
         int Dummy2();
         int Dummy3();
         int Dummy4();
      }

      [StructLayout(LayoutKind.Sequential)]
      internal struct AssemblyInfo
      {
         public int cbAssemblyInfo;
         public int assemblyFlags;
         public long assemblySizeInKB;

         [MarshalAs(UnmanagedType.LPWStr)]
         public String currentAssemblyPath;

         public int cchBuf;
      }

      /// <summary>
      /// Simple method to retrieve full path to best matching GAC library from a short name reference
      /// </summary>
      /// <param name="assemblyName">Short name of assembly to search for</param>
      /// <returns>Full path to assembly, if found</returns>
      /// <exception></exception>
      public static String QueryAssemblyInfo(string assemblyName)
      {
         var assembyInfo = new AssemblyInfo { cchBuf = 512 };
         assembyInfo.currentAssemblyPath = new String('\0', assembyInfo.cchBuf);

         IAssemblyCache assemblyCache;

         // Get IAssemblyCache pointer
         var hr = GacApi.CreateAssemblyCache(out assemblyCache, 0);
         if (hr == IntPtr.Zero)
         {
            hr = assemblyCache.QueryAssemblyInfo(1, assemblyName, ref assembyInfo);
            if (hr != IntPtr.Zero)
               Marshal.ThrowExceptionForHR((int)hr);
         }
         else
         {
            Marshal.ThrowExceptionForHR((int)hr);
         }
         return assembyInfo.currentAssemblyPath;
      }
   }
}
