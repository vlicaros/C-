using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;
using System.Diagnostics;
using System.Reflection;
using System.IO;
using System.ServiceProcess;
using System.Security.Principal;
using System.Data.SqlLocalDb;
using System.Data.SqlClient;

namespace CustomActions
{
   public partial class CustomActions
   {
      // Each key/value should be { "InstanceName", "localDBVersionNumber" }
      internal static readonly Dictionary<String, String> localDBInstances = new Dictionary<String, String> {
         {"BarTender_DataBuilder_2019", "12.0"},
         {"BarTenderLicensingDB120","12.0"},
         {"PrintSchedulerDB120", "12.0"},
         {"MaestroServiceDB120", "12.0"}
      };

      /// <summary>
      /// Adjusts property values for silent installs
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult SilentInstallProperties(Session session)
      {
         // Error out if we are given MSI quiet or basic options
         String cmdLine = session["EXE_CMD_LINE"].ToLowerInvariant();
         if (cmdLine.Contains("/qn") || cmdLine.Contains("/qb") || cmdLine.Contains("/quiet"))
         {
            session.Log("This installer does not support MSI style basic and quiet flags.  Please review silent install / commandline help.");
            return ActionResult.Failure;
         }

         // Using the FEATURE properties will set silent install mode, only process this once
         if (session["FEATURE"].Trim() != string.Empty && session["UILevel"] != "2")
         {
            // Set user interface level to silent
            session["UILevel"] = "2";

            // Validate FEATURE
            if (!MappedInstallFeatures.ContainsKey(session["FEATURE"]))
            {
               session.Log("FEATURE value \"" + session["FEATURE"] + "\" is not a permitted property value.  Please review silent install / commandline help.");
               return ActionResult.Failure;
            }
            else if (MappedInstallFeatures.ContainsKey(session["FEATURE"]))
            {
               session.Log("FEATURE value \"" + session["FEATURE"] + "\" is a permitted property value.");
            }

            ////// SET FEATURES //////

            // Parse the feature list selected by the user either via UI or commandline
            List<InstallFeatures> featureList = GetFeatures(session["FEATURE"], "");

            foreach (InstallFeatures feature in Enum.GetValues(typeof(InstallFeatures)))
            {
               try
               {
                  session.Log("Pre Marking feature {0} ", feature.ToString());
                  if (featureList.Contains(feature))
                  {
                     session.Log("{0} is contained in FeatureList ", feature.ToString());
                     session.Features[feature.ToString()].RequestState = InstallState.Local;
                     session.Log("Marking feature {0} as InstallState.Local", feature.ToString());
                  }
                  else
                  {
                     session.Log("{0} is NOT contained in FeatureList ", feature.ToString());
                     session.Features[feature.ToString()].RequestState = InstallState.Absent;
                     session.Log("Marking feature {0} as InstallState.Absent", feature.ToString());
                  }
               }
               catch (Exception ex)
               {
                  session.Log("Exception has occurred: " + ex.Message);
               }
            }
         }

         if (session["REMOVE"].Trim() != string.Empty && session["UILevel"] != "2" && session["SECONDSEQUENCE"] != "1")
         {
            session["UILevel"] = "2";
            try
            {
               if ((session["REMOVE"].ToLowerInvariant() == "all"))
               {
                  session["InstallMode"] = "Remove";
                  session["AI_INSTALL_MODE"] = "Remove";
               }
               else
               {
                   session.Log("REMOVE value \"" + session["REMOVE"] + "\" is not a permitted property value.  Please use REMOVE=ALL commandline.");
                  return ActionResult.Failure;
               }
            }
            catch (Exception ex)
            {
               session.Log("Exception has occurred: " + ex.Message);
            }
         }

         // Santize INSTALLDIR if set and populate into APPDIR/TARGETDIR as well
         if (session["INSTALLDIR"].Trim() != string.Empty)
         {
            string sep = System.IO.Path.DirectorySeparatorChar.ToString();
            session["INSTALLDIR"] = session["INSTALLDIR"] + ((session["INSTALLDIR"].EndsWith(sep)) ? "" : sep);
            session["APPDIR"] = session["INSTALLDIR"];
            session["TARGETDIR"] = session["INSTALLDIR"];
         }

         // Santize the PKC property if one was provided
         if (session["PKC"].Trim() != string.Empty)
         {
            // Santize PKC to XXXX-XXXX-XXXX-XXXX
            string pkc = System.Text.RegularExpressions.Regex.Replace(session["PKC"].Trim(), "[^A-Za-z0-9]", "");
            if (pkc.Length != 16)
            {
               session.Log("PKC property must be 19 characters (with hyphens) or 16 characters (without hyphens) long.  Please review silent install / commandline help.");
               return ActionResult.Failure;
            }
            List<string> fixedPKC = new List<string> { pkc.Substring(0, 4).ToUpper(),
                                                       pkc.Substring(4, 4).ToUpper(),
                                                       pkc.Substring(8, 4).ToUpper(),
                                                       pkc.Substring(12, 4).ToUpper() };
            session["PKC"] = String.Join("-", fixedPKC);
         }

         // Sanitize the BLS property if one was provided
         if (session["BLS"].Trim() != string.Empty)
         {
            string[] uri = session["BLS"].Split(':');
            if (uri.Length == 0 || uri.Length > 2)
            {
               session.Log("BLS property must be format as either HOST or HOST:PORT.  Please review silent install / commandline help.");
               return ActionResult.Failure;
            }
            int port = 5160;
            if (uri.Length > 1)
               Int32.TryParse(uri[1], out port);
            session["BLS"] = uri[0] + ":" + port.ToString();
         }

         // Ensure that if user selects PrintPortal that we also have a PRINTPORTAL_ACCOUNT_PASSWORD
         if (session["PRINTPORTAL_ACCOUNT_PASSWORD"].Trim() == string.Empty && (session["FEATURE"].ToLowerInvariant() == "printportal"))
         {
            session.Log("PRINTPORTAL_ACCOUNT_PASSWORD is required when installing PrintPortal.  Please review silent install / commandline help.");
            return ActionResult.Failure;
         }

         // Check for user controlled INSTALLSQL property
         if (session["INSTALLSQL"].Trim() != string.Empty)
         {
            session["INSTALLSQL"] = (session["INSTALLSQL"].ToLowerInvariant() != "true") ? "false" : "true";
         }

         // Check for Windows Update reboot required
         CustomActions.GetWindowsUpdateRebootRequired(session);
         if (session["WindowsUpdateRebootRequired"] == "true")
         {
            session.Log("Windows Update has scheduled a system restart. Please restart your system then re-launch BarTender Setup.");
            return ActionResult.Failure;
         }

         return ActionResult.Success;
      }

      /// <summary>
      /// Performs deactivation steps during install execute sequence when removing product only
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult DoDeactivation(Session session)
      {
         // Check for DEACTIVATE during uninstall, but only during second sequence
         if (session["DEACTIVATE"].ToLowerInvariant() == "true"
              && session["SECONDSEQUENCE"].Trim() != string.Empty
              && (session["REMOVE"].ToLowerInvariant() == "all" || session["AI_INSTALL_MODE"].ToLowerInvariant() == "remove"))
         {
            //ActivationWizard.exe Deactivate
            try
            {
               // Begin preparing commandline
               System.Diagnostics.Process prc = new System.Diagnostics.Process();
               prc.StartInfo.UseShellExecute = false;
               prc.StartInfo.FileName = "cmd.exe";
               prc.StartInfo.Verb = "runas";
               prc.StartInfo.RedirectStandardOutput = true;
               prc.StartInfo.RedirectStandardError = true;
               prc.StartInfo.CreateNoWindow = true;
               prc.StartInfo.WorkingDirectory = session["INSTALLDIR"];
               prc.StartInfo.Arguments = @"/C """"" + System.IO.Path.Combine(session["INSTALLDIR"], "ActivationWizard.exe") + @""" Deactivate";

               // Finish arguments and execute
               prc.StartInfo.Arguments += "\"";
               session.Log("Running: " + prc.StartInfo.Arguments);
               prc.Start();
            }
            catch (Exception ex)
            {
               session.Log("Exception running ActivationWizard: " + ex.Message);
            }
         }
         return ActionResult.Success;
      }

      /// <summary>
      /// Combines all upgradecodes for incompatible older BarTender installations, setting OLDPRODUCTS
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult ForceUpgradeProperty(Session session)
      {
         List<string> upgradeCodes = new List<string>();
         string newUpgradeCodeList = string.Empty;

         try
         {
            // UpgradeCodes we do want to 'replace'
            upgradeCodes.Add(session["OLDPRODUCTS"]);
            upgradeCodes.Add(session["UPGRADECODE_BT_CURRENT"]);
            upgradeCodes.Add(session["UPGRADECODE_SLS_90_CURRENT"]);
            upgradeCodes.Add(session["UPGRADECODE_801_2179_UL"]);
            upgradeCodes.Add(session["UPGRADECODE_801_2179"]);
            upgradeCodes.Add(session["UPGRADECODE_SLS_801_2179"]);
            upgradeCodes.Add(session["UPGRADECODE_801_2275"]);
            upgradeCodes.Add(session["UPGRADECODE_Rimage_801"]);
            upgradeCodes.Add(session["UPGRADECODE_Hybrid_80"]);
            upgradeCodes.Add(session["UPGRADECODE_UL_CURRENT"]);

            // Upgradecodes we don't want to touch yet...
            //upgradeCodes.Add(session["UPGRADECODE_RIMAGE_CURRENT"]);
         }
         catch (Exception ex) { session.Log(ex.Message); }

         foreach (string ucode in upgradeCodes)
         {
            if (!string.IsNullOrEmpty(ucode))
            {
               try
               {
                  Guid gu = new Guid(ucode);
                  foreach (ProductInstallation prod in ProductInstallation.GetRelatedProducts(gu.ToString("B")))
                  {
                     newUpgradeCodeList += prod.ProductCode + ";";
                  }
               }
               catch (Exception ex) { session.Log(ex.Message); }
            }
         }
         // If we found at least one GUID...
         if (newUpgradeCodeList.Length > 36)
            session["OLDPRODUCTS"] = newUpgradeCodeList;

         return ActionResult.Success;
      }

      /// <summary>
      /// Determines version of any existing version of BarTender installed already
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult SetInstalledVersion(Session session)
      {
         string upgradeCode = session["OLDPRODUCTS"];
         if (upgradeCode.Contains(";"))
         {
            upgradeCode = upgradeCode.Split(";".ToCharArray())[0];
         }

         bool hasInstalledSession = false;
         Session installedSession = null;
         try
         {
            installedSession = Installer.OpenProduct(upgradeCode);
         }
         catch (Exception ex)
         {
            session.Log(String.Format("Error attempting to detect previous installed products: {0}", ex.Message));
         }
         finally
         {
            if (installedSession != null)
            {
               session.Log(String.Format("Matching product installation was found with UpgradeCode: {0}", upgradeCode));
               hasInstalledSession = true;
            }
            else
            {
               session.Log(String.Format("No matching product was found"));
            }
         }

         // If previous installed session detected, then create a view to the installer database and execute query for version.
         Microsoft.Deployment.WindowsInstaller.View InstalledVersionView = null;
         string installedVersion = string.Empty;
         if (hasInstalledSession)
         {
            try
            {
               // post 11.1 - use FullProductVersion (4 digit version number)
               InstalledVersionView = installedSession.Database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'FullProductVersion'");
               InstalledVersionView.Execute();
               Record record = InstalledVersionView.Fetch();
               installedVersion = record[1].ToString();
            }
            catch { }
            finally
            {
               if (string.IsNullOrEmpty(installedVersion))
               {
                  try
                  {
                     // pre 11.1 - use ProductVersion (3 digit version number)
                     InstalledVersionView = installedSession.Database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'ProductVersion'");
                     InstalledVersionView.Execute();
                     Record record = InstalledVersionView.Fetch();
                     installedVersion = record[1].ToString();
                  }
                  catch (Exception ex)
                  {
                     session.Log(String.Format("Error reading database of previously installed product: {0}", ex.Message));
                  }
               }

               if (!string.IsNullOrEmpty(installedVersion))
               {
                  session["BT_INSTALLED_VERSION"] = installedVersion;
               }
            }
         }

         bool trySlsVersion = false;

         // If previous installed session detected, then create a view to the installer database and execute query for version.
         Microsoft.Deployment.WindowsInstaller.View InstalledVersionStringView = null;
         string installedVersionString = string.Empty;
         if (hasInstalledSession)
         {
            try
            {
               InstalledVersionStringView = installedSession.Database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'BT_PRODUCT_VERSION'");
               InstalledVersionStringView.Execute();
               Record record = InstalledVersionStringView.Fetch();
               installedVersionString = record[1].ToString();
            }
            catch
            {
               trySlsVersion = true;
            }
            finally
            {
               if (trySlsVersion)
               {
                  try
                  {
                     InstalledVersionStringView = installedSession.Database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'LS_PRODUCT_VERSION'");
                     InstalledVersionStringView.Execute();
                     Record record = InstalledVersionStringView.Fetch();
                     installedVersionString = record[1].ToString();
                  }
                  catch (Exception ex)
                  {
                     session.Log(String.Format("Error reading database of previously installed product: {0}", ex.Message));
                  }
               }

               if (!string.IsNullOrEmpty(installedVersion))
               {
                  session["BT_INSTALLED_VERSION_STRING"] = installedVersionString;
               }
            }
         }

         // Lookup the old installers install location, if possible
         if (hasInstalledSession)
         {
            try
            {
               RegistryKey rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
               using (RegistryKey key = rk.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + upgradeCode))
               {
                  if (key != null)
                     session["BT_INSTALLED_INSTALLLOCATION"] = key.GetValue("InstallLocation", "").ToString();
               }
               using (RegistryKey key = rk.OpenSubKey(@"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\" + upgradeCode))
               {
                  if (key != null)
                     session["BT_INSTALLED_INSTALLLOCATION"] = key.GetValue("InstallLocation", "").ToString();
               }
               rk.Dispose();
            }
            catch (Exception e)
            {
               session.Log("Error obtaining InstallLocation from registry: " + e.Message);
            }
         }

         // Close the installed session
         if (hasInstalledSession && installedSession != null && !installedSession.IsClosed)
            installedSession.Dispose();

         // START WORKAROUND - SIGMA-2871

         // Move items from OLDPRODUCTS to BETAPRODUCTS where we include extra arguments to work around a bug
         // introduced in installer's upgrade behavior.  Only affects builds between ranges listed below
         // and should be removed at a future date when the listed range is highly improbable to still exist
         // in the wild
         if (session["OLDPRODUCTS"].Trim() != string.Empty && session["BETAPRODUCTS"].Trim() != string.Empty)
         {
            try
            {
               string oldProductCodes = session["OLDPRODUCTS"].Trim();
               string betaProductCodes = session["BETAPRODUCTS"].Trim();
               Version betaProductLow = new Version("11.2.0.151957");
               Version betaProductHigh = new Version("11.2.0.154200");

               string allKnownProductCodes = session["OLDPRODUCTS"].Trim() + ";" + session["BETAPRODUCTS"].Trim();

               foreach (string currentProductCode in allKnownProductCodes.Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Distinct())
               {
                  // Continue if not a valid ProductCode
                  if (currentProductCode.Length != 38) continue;

                  try
                  {
                     // Get FullProductVersion from ProductCode
                     Session currentProductSession = Installer.OpenProduct(currentProductCode);
                     View currentProductView = currentProductSession.Database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = 'FullProductVersion'");
                     currentProductView.Execute();
                     Record record = currentProductView.Fetch();
                     Version currentProductVersion = new Version(record[1].ToString());
                     currentProductSession.Dispose();

                     if (currentProductVersion >= betaProductLow && currentProductVersion <= betaProductHigh)
                     {
                        // It's in our range, add to beta and remove from old
                        session.Log("Sorting to BETAPRODUCTS: " + currentProductCode);
                        betaProductCodes = String.Join(";", (betaProductCodes + ";" + currentProductCode).Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Distinct());
                        oldProductCodes = String.Join(";", oldProductCodes.Replace(currentProductCode, "").Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Distinct());
                     }
                     else
                     {
                        // It's not in our range, remove from beta and add to old
                        session.Log("Sorting to OLDPRODUCTS: " + currentProductCode);
                        oldProductCodes = String.Join(";", (oldProductCodes + ";" + currentProductCode).Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Distinct());
                        betaProductCodes = String.Join(";", betaProductCodes.Replace(currentProductCode, "").Split(";".ToCharArray(), StringSplitOptions.RemoveEmptyEntries).Distinct());
                     }
                  }
                  catch (Exception e)
                  {
                     session.Log("Exception in BETAPRODUCTS/OLDPRODUCTS: " + e.Message);
                  }
               }

               // Update OLDPRODUCTS and BETAPRODUCTS
               session["OLDPRODUCTS"] = oldProductCodes;
               session["BETAPRODUCTS"] = betaProductCodes;
            }
            catch { }
         }
         // END WORKAROUND - SIGMA-2871

         // No matter what the results, we want to return Sucess
         return ActionResult.Success;
      }

      /// <summary>
      /// Enables all required windows optional features
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult WindowsOptionalFeatures(Session session)
      {
         // Operating system info
         bool is_server = IsOS(OS_ANYSERVER);
         Version win8 = new Version(6, 2);
         Version currentOS = new Version(Environment.OSVersion.Version.Major, Environment.OSVersion.Version.Minor);

         // Get current DISM features
         Dictionary<string, string> currentFeatures = CustomActions.GetCurrentWindowsOptionalFeatures();

         // Install MSMQ (always)
         List<string> MSMQ_features;
         if (is_server)
            MSMQ_features = new List<string>() { "MSMQ", "MSMQ-Server" };
         else
            MSMQ_features = new List<string>() { "MSMQ-Container", "MSMQ-Server" };

         if (!CustomActions.EnableWindowsOptionalFeatures(session, MSMQ_features, currentFeatures))
            session.Log("Encountered problem enabling Windows optional features for MSMQ support");

         // Install IIS (if BPP was selected)
         if (session.Features["BPP"].RequestState == InstallState.Local)
         {
            List<string> IIS_features;
            if (currentOS >= win8)
               IIS_features = new List<string>() { "IIS-ASPNET45", "IIS-NetFxExtensibility45", "IIS-DefaultDocument", "IIS-ISAPIExtensions", "IIS-ISAPIFilter", "IIS-WebServer", "IIS-WebServerRole", "IIS-WebSockets" };
            else
               IIS_features = new List<string>() { "IIS-ASPNET", "IIS-NetFxExtensibility", "IIS-DefaultDocument", "IIS-ISAPIExtensions", "IIS-ISAPIFilter", "IIS-WebServer", "IIS-WebServerRole" };

            // Common features (from AdvancedInstaller's default IIS configuration)
            /*IIS_features.AddRange(new List<string>() { "IIS-ServerSideIncludes", "IIS-RequestFiltering", "IIS-WebServerRole", "IIS-DefaultDocument",
                                                       "IIS-HttpRedirect", "IIS-ISAPIFilter", "IIS-ASP", "IIS-CGI", "IIS-ISAPIExtensions", "IIS-DirectoryBrowsing",
                                                       "IIS-HttpErrors", "IIS-ODBCLogging", "IIS-StaticContent", "IIS-WebDAV", "IIS-CustomLogging", "IIS-HttpLogging",
                                                       "IIS-HttpTracing", "IIS-BasicAuthentication", "IIS-LoggingLibraries", "IIS-RequestMonitor", "IIS-HttpCompressionDynamic",
                                                       "IIS-HttpCompressionStatic", "IIS-ClientCertificateMappingAuthentication", "IIS-DigestAuthentication",
                                                       "IIS-IISCertificateMappingAuthentication", "IIS-IPSecurity", "IIS-RequestFiltering", "IIS-IIS6ManagementCompatibility",
                                                       "IIS-Metabase", ""IIS-WebServerManagementTools", "IIS-URLAuthorization", "IIS-WindowsAuthentication",
                                                       "IIS-ManagementScriptingTools" });
             */
            if (!CustomActions.EnableWindowsOptionalFeatures(session, IIS_features, currentFeatures))
               session.Log("Encountered problem enabling Windows optional features for IIS support");
            else
            {
               // Update installer's IIS version property after installation
               session["IIS_VERSION"] = IISVersion();
               if (session["IIS_VERSION"] == "")
                  session.Log("Unable to get IIS version");
            }

            // Install IIS Rewrite package
            try
            {
               RegistryKey rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
               string registryPath = String.Format(@"SOFTWARE\Microsoft\IIS Extensions\URL Rewrite\Install");
               using (RegistryKey key = rk.OpenSubKey(registryPath))
               {
                  if (key == null)
                  {
                     if (Environment.Is64BitOperatingSystem)
                     {
                        if (CustomActions.ExtractIISPrereqInstallerBinary(session, "IISRewriteInstallerBinaryamd64", session["IISREWRITEAMD64"]))
                           CustomActions.InstallIISPrerequisiteBinary(session, session["IISREWRITEAMD64"]);
                        else
                           session.Log("Encountered problem installing IIS Rewrite 64-bit module ");
                     }
                     else
                     {
                        if (CustomActions.ExtractIISPrereqInstallerBinary(session, "IISRewriteInstallerBinaryx86", session["IISREWRITEX86"]))
                           CustomActions.InstallIISPrerequisiteBinary(session, session["IISREWRITEX86"]);
                        else
                           session.Log("Encountered problem installing IIS Rewrite 32-bit module");

                     }
                  }
               }
               rk.Dispose();
            }
            catch (Exception e)
            {
               session.Log("Exception: " + e.Message);
               return ActionResult.Failure;
            }

            // Install IIS Application Request Routing package
            try
            {
               RegistryKey rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
               string registryPath = String.Format(@"SOFTWARE\Microsoft\IIS Extensions\Application Request Routing\Install");
               using (RegistryKey key = rk.OpenSubKey(registryPath))
               {
                  if (key == null)
                  {
                     if (Environment.Is64BitOperatingSystem)
                     {
                        if (CustomActions.ExtractIISPrereqInstallerBinary(session, "IISApplicationRequestRouterInstallerBinaryamd64", session["IISREQUESTROUTERAMD64"]))
                           CustomActions.InstallIISPrerequisiteBinary(session, session["IISREQUESTROUTERAMD64"]);
                        else
                           session.Log("Encountered problem installing IIS Application Request Router 64-bit module ");
                     }
                     else
                     {
                        if (CustomActions.ExtractIISPrereqInstallerBinary(session, "IISApplicationRequestRouterInstallerBinaryx86", session["IISREQUESTROUTERX86"]))
                           CustomActions.InstallIISPrerequisiteBinary(session, session["IISREQUESTROUTERX86"]);
                        else
                           session.Log("Encountered problem installing IIS Application Request Router 32-bit module");
                     }
                  }
               }
               rk.Dispose();
            }
            catch (Exception e)
            {
               session.Log("Exception: " + e.Message);
               return ActionResult.Failure;
            }
         }


         // Install NetFX3 (if MSSQL is to be installed)
         if (session["INSTALLSQL"] == "true")
         {
            /*
            List<string> NetFx3_features = new List<string>() { "NetFx3" };

            if (!CustomActions.EnableWindowsOptionalFeatures(session, NetFx3_features, currentFeatures))
               session.Log("Encountered problem enabling Windows optional features for NetFx3 support");
             */
         }

         // Enable NetTcpSharing service (always)
         String serviceModelRegPath = "";
         if (Environment.Is64BitOperatingSystem)
            serviceModelRegPath = Environment.GetEnvironmentVariable("windir") + @"\Microsoft.Net\Framework64\v4.0.30319\ServiceModelReg.exe";
         else
            serviceModelRegPath = Environment.GetEnvironmentVariable("windir") + @"\Microsoft.Net\Framework\v4.0.30319\ServiceModelReg.exe";
         String output = CustomActions.RunCommandAndCaptureAllOutputs(serviceModelRegPath, "-r");
         if (output.ToLowerInvariant().Contains("[error]"))
            session.Log("Error in ServiceModelReg: " + output);

         // Always return success
         return ActionResult.Success;
      }

      /// <summary>
      /// Determines current status of all windows optional features on this machine
      /// </summary>
      /// <param></param>
      /// <returns>Dictionary of feature name with status text</returns>
      public static Dictionary<string, string> GetCurrentWindowsOptionalFeatures()
      {
         Dictionary<string, string> retVal = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
         Regex pattern = new Regex(@"^(?<Feature>\S+)\s+\|\s(?<Status>.*)$");

         try
         {
            // Call "dism.exe /Online /Get-Features /Format:Table"
            String dismPath;
            if ((Environment.Is64BitProcess == false) && (Environment.Is64BitOperatingSystem == true))
               dismPath = Environment.GetEnvironmentVariable("windir") + @"\sysnative\dism.exe";
            else
               dismPath = Environment.SystemDirectory + @"\dism.exe";
            String output = CustomActions.RunCommandAndCaptureAllOutputs(dismPath, "/Online /Get-Features /Format:Table");

            foreach (var line in output.Split(System.Environment.NewLine.ToArray()))
            {
               // Ignore lines starting with -----
               if (line.StartsWith("-----"))
                  continue;

               MatchCollection matches = pattern.Matches(line);
               foreach (Match match in matches)
               {
                  GroupCollection groups = match.Groups;
                  retVal[groups["Feature"].Value.Trim()] = groups["Status"].Value.Trim();
               }
            }
         }
         catch (Exception)
         {
            // Ignore
         }
         return retVal;
      }

      /// <summary>
      /// Runs dism.exe to add requested features, provided they are not already enabled and they exist
      /// </summary>
      /// <param>session, list of requested features, and dictionary of current feature states</param>
      /// <returns>boolean success/failure</returns>
      public static bool EnableWindowsOptionalFeatures(Session session, List<string> requestedFeatures, Dictionary<string, string> currentFeatures)
      {
         try
         {
            String dismPath;
            String arguments;
            Version win8 = new Version(6, 2);
            Version currentOS = new Version(Environment.OSVersion.Version.Major, Environment.OSVersion.Version.Minor);
            bool currentFeaturesAreUsable = (currentFeatures != null && currentFeatures.Count > 20 && currentFeatures.ContainsValue("Enabled"));

            // Create a filtered list of requested features
            List<string> requestedFeaturesFiltered = requestedFeatures.Distinct(StringComparer.InvariantCultureIgnoreCase).ToList();
            if (currentFeaturesAreUsable)
            {
               requestedFeaturesFiltered.RemoveAll(x => !currentFeatures.ContainsKey(x));       // Remove non-existant features
               requestedFeaturesFiltered.RemoveAll(x => currentFeatures[x].Contains("Enable")); // Remove already enabled features
            }
            else
            {
               // Do no filter if currentFeatures are not 'usable'
            }

            // Return now if no work needs to be done
            if (requestedFeaturesFiltered.Count == 0)
            {
               session.Log("Requested feature(s) are already enabled or not available on this system");
               return true;
            }

            // Set path to dism.exe
            if ((Environment.Is64BitProcess == false) && (Environment.Is64BitOperatingSystem == true))
               dismPath = Environment.GetEnvironmentVariable("windir") + @"\sysnative\dism.exe";
            else
               dismPath = Environment.SystemDirectory + @"\dism.exe";

            // Format our arguments
            arguments = @"/Online /Enable-Feature ";
            foreach (String request in requestedFeaturesFiltered)
            {
               // Sanitize requested feature
               String adjustedFeatureName = request.Trim();
               if (currentFeaturesAreUsable)
                  adjustedFeatureName = currentFeatures.FirstOrDefault(kvp => kvp.Key.ToLowerInvariant() == request.Trim().ToLowerInvariant()).Key;
               // Add the requested feature
               arguments += @"/FeatureName:" + adjustedFeatureName + @" ";
            }
            // Win8+ should always add an '/All' argument
            if (currentOS >= win8)
               arguments += @"/All ";
            arguments += @"/NoRestart";

            // Run dism.exe
            session.Log("Running: " + dismPath + " " + arguments);
            String output = CustomActions.RunCommandAndCaptureAllOutputs(dismPath, arguments);

            // Parse output
            if (output.ToLowerInvariant().Contains("completed successfully"))
            {
               session.Log("Requested feature(s) were successfully enabled");
               return true;
            }
            else
            {
               session.Log("Dism error: " + output);
               return false;
            }
         }
         catch (Exception e)
         {
            session.Log("Exception in EnableWindowsOptionalFeatures: " + e.ToString());
         }
         return false;
      }

      public static ActionResult InstallIISPrerequisiteBinary(Session session, String fileName)
      {
         //run MSI
         // Begin preparing commandline
         System.Diagnostics.Process prc = new System.Diagnostics.Process();
         prc.StartInfo.UseShellExecute = false;
         prc.StartInfo.Verb = "runas";
         prc.StartInfo.RedirectStandardOutput = true;
         prc.StartInfo.RedirectStandardError = true;
         prc.StartInfo.CreateNoWindow = true;
         prc.StartInfo.WorkingDirectory = session["TempFolder"];
         prc.StartInfo.FileName = "cmd.exe";
         prc.StartInfo.Arguments = @"/C """ + @"msiexec.exe /i " + System.IO.Path.Combine(session["TempFolder"], fileName) +
         @" /q /l*v C:\ProgramData\Seagull\Installer\IIS_" + fileName + ".log";

         // Finish arguments and execute
         prc.StartInfo.Arguments += "\"";
         session.Log("Running: " + prc.StartInfo.Arguments);
         try
         {
            prc.Start();
            prc.WaitForExit();

            // Log any errors
            if (prc.ExitCode != 0 && prc.ExitCode != 1618)
            {
               session.Log("Error during install of " + fileName + ", exit code was " + prc.ExitCode);
               return ActionResult.Failure;
            }
         }
         catch (Exception ex)
         {
            session.Log("Exception running " + fileName + ": " + ex.Message);
            return ActionResult.Failure;
         }
         return ActionResult.Success;
      }
      /// <summary>
      /// This will open the MSI database and extract the IIS Application Request Router binary from the binary table
      /// into the users temp folder
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success if extraction was completed. ActionResult.Failure if there is a problem</returns>
      [CustomAction]
      public static bool ExtractIISPrereqInstallerBinary(Session session, string binaryName, string fileName)
      {
         // Execute query on MSI database for the IISReRouterInstallerBinary
         View dbView = session.Database.OpenView("SELECT * FROM Binary WHERE Name = '" + binaryName + "'");
         dbView.Execute();

         // Fetch first record (assume one), get the "Data" field (index 2)
         Record rec = dbView.Fetch();
         Stream prereqBinary = rec.GetStream("Data");

         // If no data or data sizes don't match expected values...
         if (!(rec.GetDataSize("Data") > 0) || rec.GetDataSize("Data") != prereqBinary.Length)
            throw new Exception("Bad or no Data field found in record from msi's Binary table");

         String outputFile = session["TempFolder"] + fileName;
         if (File.Exists(outputFile))
            File.Delete(outputFile);

         // Copy data to outputfile
         using (FileStream fs = File.Create(outputFile))
         {
            prereqBinary.CopyTo(fs);
         }
         session.Log(binaryName + " extracted to " + outputFile);
         return true;
      }

      /// <summary>
      /// Creates a backup of the user's BarTender documents
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult BackupDocuments(Session session)
      {
         string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
         session.Log(Path.Combine(myDocs, "BarTender"));
         if (Directory.Exists(Path.Combine(myDocs, "BarTender")))
         {
            try
            {
               if (!Directory.Exists(Path.Combine(myDocs, "BarTender_OLD")))
               {
                  Directory.CreateDirectory(Path.Combine(myDocs, "BarTender_OLD"));
                  Microsoft.VisualBasic.Devices.Computer comp = new Microsoft.VisualBasic.Devices.Computer();
                  comp.FileSystem.CopyDirectory(Path.Combine(myDocs, "BarTender"), Path.Combine(myDocs, "BarTender_OLD"));
               }
            }
            catch (Exception ex)
            {
               session.Log(ex.Message);
            }
         }
         return ActionResult.Success;
      }

      /// <summary>
      /// Restores the users BarTender documents after default documents are added
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult RestoreDocuments(Session session)
      {
         try
         {
            string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            session.Log(Path.Combine(myDocs, "BarTender_OLD"));
            if (Directory.Exists(Path.Combine(myDocs, "BarTender_OLD")))
            {
               session.Log("Copying files");
               Microsoft.VisualBasic.Devices.Computer comp = new Microsoft.VisualBasic.Devices.Computer();
               comp.FileSystem.CopyDirectory(Path.Combine(myDocs, "BarTender_OLD"), Path.Combine(myDocs, "BarTender"), true);
            }
         }

         catch (Exception ex)
         {
            session.Log(ex.Message);
         }

         return ActionResult.Success;
      }

      /// <summary>
      /// Cleans up leftovers from running the BackupDocuments and RestoreDocuments CAs
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult RestoreFormatsCleanup(Session session)
      {
         string[] badFiles = { string.Empty };
         if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender_OLD")))
         {
            badFiles = Directory.GetFiles(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender_OLD"), "blank.txt", SearchOption.AllDirectories);
         }
         foreach (string fl in badFiles)
         {
            try
            {
               FileInfo fileinf = new FileInfo(fl);
               fileinf.Attributes = System.IO.FileAttributes.Normal;
               File.Delete(fl);
            }
            catch (Exception ex) { session.Log(ex.Message); }
         }

         try
         {
            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender_OLD")))
            {
               Directory.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender_OLD"), true);
            }
         }
         catch (Exception ex)
         {
            session.Log(ex.Message);
         }

         // Only gets set to 0 if user has selected to remove formats
         string[] badFiles2 = { string.Empty };
         if (session["BACKUP_FORMATS"] == "0")
         {
            if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender")))
            {
               badFiles2 = Directory.GetFiles(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender"), "blank.txt", SearchOption.AllDirectories);
            }
            foreach (string fl in badFiles2)
            {
               try
               {
                  FileInfo fileinf = new FileInfo(fl);
                  fileinf.Attributes = System.IO.FileAttributes.Normal;
                  File.Delete(fl);
               }
               catch (Exception ex) { session.Log(ex.Message); }
            }

            try
            {
               if (Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender")))
               {
                  Directory.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BarTender"), true);
               }
            }
            catch (Exception ex)
            {
               session.Log(ex.Message);
            }
         }

         return ActionResult.Success;
      }

      /// <summary>
      /// SetBootStrapperLanguage gets the preferred user language to set the BootStrapper Language.
      /// Currently, it will make specific custom logic for Spanish.
      /// Registry key settings for HKCU and HKU are not accessible in BootstrapperUI sequence
      /// Using PowerShell command to obtain the list of Windows User Preferred languages.
      /// Readline will get the default language currently in use and detect if it is a variant of "es"
      /// language tag
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult SetBootStrapperLanguage(Session session)
      {

        String sWinUserLang = string.Empty;
        System.Diagnostics.Process process = new System.Diagnostics.Process();
        process.StartInfo.UseShellExecute = false;
        process.StartInfo.Verb = "Open";
        process.StartInfo.RedirectStandardOutput = true;
        process.StartInfo.RedirectStandardError = false;
        process.StartInfo.CreateNoWindow = true;
        process.StartInfo.WorkingDirectory = session["TempFolder"];
        process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
        //Use PowerShell command to get default langauge currently in use
        process.StartInfo.FileName = "powershell.exe";
        process.StartInfo.Arguments = @"(Get-WinUserLanguageList).LanguageTag";

        // Execute and wait
        session.Log("Get default Windows User Language with PowerShell: " + process.StartInfo.Arguments);
        process.Start();
        // Get the first language in the list
        sWinUserLang = process.StandardOutput.ReadLine();
        session.Log("Default Current User Language is :" + sWinUserLang);
        process.WaitForExit();
        
        if (sWinUserLang.StartsWith("es"))
        {
            session["AI_BOOTSTRAPPERLANG"] = "3082";
            session["TRANSFORMS"] = "3082";
        }
         return ActionResult.Success;
      }

         /// <summary>
         /// SetApplicationsLanguage sets the registry key for application language
         /// HKLM\\SOFTWARE\\Seagull Scientific\\Common\\Preferences -> Language
         /// </summary>
         /// <param name="session"></param>
         /// <returns></returns>
         [CustomAction]
      public static ActionResult SetApplicationsLanguage(Session session)
      {
         // Check reg key doesn't already exist
         RegistryKey regLocalMachine = Registry.LocalMachine;
         RegistryKey regCommonPrefences = null;
         String sRegLanguage = string.Empty;
         try
         {
            regCommonPrefences = regLocalMachine.OpenSubKey("\\SOFTWARE\\Seagull Scientific\\Common\\Preferences");
            sRegLanguage = (String)regCommonPrefences.GetValue("Language", String.Empty);
         }
         catch { }

         string sInstallLanguage = string.Empty;
         switch (session["ProductLanguage"])
         {
            case "2052":
               sInstallLanguage = "zh-CN";
               break;
            case "1028":
               sInstallLanguage = "zh-TW";
               break;
            case "1029":
               sInstallLanguage = "cs-CZ";
               break;
            case "1030":
               sInstallLanguage = "da-DK";
               break;
            case "1031":
               sInstallLanguage = "de-DE";
               break;
            case "1033":
               sInstallLanguage = string.Empty;
               break;
            case "1034":
            case "3082":
               sInstallLanguage = "es-ES";
               break;
            case "1035":
               sInstallLanguage = "fi-FI";
               break;
            case "1036":
               sInstallLanguage = "fr-FR";
               break;
            case "1040":
               sInstallLanguage = "it-IT";
               break;
            case "1041":
               sInstallLanguage = "jp-JP";
               break;
            case "1042":
               sInstallLanguage = "ko-KR";
               break;
            case "1043":
               sInstallLanguage = "nl-NL";
               break;
            case "1044":
               sInstallLanguage = "nb-NO";
               break;
            case "1045":
               sInstallLanguage = "pl-PL";
               break;
            case "1046":
               sInstallLanguage = "pt-BR";
               break;
            case "2070":
               sInstallLanguage = "pt-PT";
               break;
            case "1049":
               sInstallLanguage = "ru-RU";
               break;
            case "1053":
               sInstallLanguage = "sv-SE";
               break;
            case "1054":
               sInstallLanguage = "th-TH";
               break;
            case "1090":
            case "1055":
               sInstallLanguage = "tr-TR";
               break;
            default:
               sInstallLanguage = string.Empty;
               break;
         }

         // Write value if it does not already have a value and the installer language is not English
         if (sRegLanguage == string.Empty && sInstallLanguage != string.Empty)
         {
            try
            {
               regCommonPrefences = regLocalMachine.OpenSubKey("\\SOFTWARE\\Seagull Scientific\\Common\\Preferences", true);
               regCommonPrefences.SetValue("Language", sInstallLanguage, RegistryValueKind.String);
            }
            catch { }
         }
         return ActionResult.Success;
      }

      /// <summary>
      /// Running NGEN on all applications after they have been registered. This should help speed up the startup of each application.
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult NGEN_Applications(Session session)
      {
         // Get release type
         string releaseBitness = session["x86_or_x64"];
         string ngenPath = string.Empty;

         try
         {
            // If 32-bit installer
            if (releaseBitness.Contains("x86"))
               ngenPath = Path.Combine(Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), @"Windows\Microsoft.NET\Framework\v4.0.30319"), "ngen.exe");
            else
               ngenPath = Path.Combine(Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), @"Windows\Microsoft.NET\Framework64\v4.0.30319"), "ngen.exe");
         }
         catch (Exception ex)
         {
            session.Log("Error:" + ex.Message);
         }
         finally
         {
            session.Log("ngen path: " + ngenPath);
         }

         if (File.Exists(ngenPath))
         {
            string[] targetAssemblies = { "bartend.exe", "activationwizard.exe", "Processbuilder.exe", "cddesign.exe", "btsystem.service.exe", "systemdatabasewizard.exe", "integrationbuilder.exe", "Integration.Service.exe", "PrintScheduler.Service.exe", "historyexplorer.exe", "librarian.exe", "maestro.exe", "maestro.service.exe", "printstation.exe", "reprintconsole.exe", "adminconsole.exe", "licensing.service.exe" };

            // Need to check if elevated, during modify from control panel installer will not be elevated
            bool isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            string installdir = (session == null ? @"C:\BarTender\Main\bin\Release-win32\" : session["INSTALLDIR"]);

            string ngenAppsCmd = "";

            // Generic process for upcoming command (SIGMA-1880)
            System.Diagnostics.Process process = new System.Diagnostics.Process();

            // Go through each app in the targetAssemblies array and create an argument string that will be appended on the previous one
            foreach (string target in targetAssemblies)
            {
               if (File.Exists(Path.Combine(installdir, target)))
               {
                  // Ampersand is only added to the argument string before the second application and so on
                  if (!target.Equals(targetAssemblies[0]))
                  {
                     ngenAppsCmd += " & ";
                  }
                  if (isAdmin)
                  {
                     session.Log("Running Ngen with admin on " + target);
                     process.StartInfo.UseShellExecute = false;
                     process.StartInfo.RedirectStandardOutput = true;
                     process.StartInfo.RedirectStandardError = true;
                     process.StartInfo.CreateNoWindow = true;
                     process.StartInfo.WorkingDirectory = installdir;

                     string ngenSingleCmd = ngenPath + " install \"" + installdir + target + "\"";
                     ngenAppsCmd += ngenSingleCmd;
                  }
                  else
                  {
                     session.Log("Running Ngen with out Admin on " + target);
                     process.StartInfo.UseShellExecute = true;
                     process.StartInfo.Verb = "runas";
                     process.StartInfo.RedirectStandardOutput = false;
                     process.StartInfo.RedirectStandardError = false;
                     process.StartInfo.CreateNoWindow = true;
                     process.StartInfo.WorkingDirectory = installdir;

                     string ngenSingleCmd = ngenPath + " install \"" + installdir + target + "\"";
                     ngenAppsCmd += ngenSingleCmd;
                  }
               }
            }
            // Log command line that will be run
            session.Log(@"cmd.exe /C " + ngenAppsCmd);
            // Start command line process
            process.StartInfo.FileName = "cmd.exe";
            process.StartInfo.Arguments = "/C " + ngenAppsCmd;
            process.Start();
         }
         else
         {
            session.Log("Could not find ngen.exe");
         }
         return ActionResult.Success;
      }

      /// <summary>
      /// This will open the MSI database and extract the SQLExpress binary from the binary table
      /// into the users temp folder
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success if extraction was completed. ActionResult.Failure if there is a problem</returns>
      [CustomAction]
      public static ActionResult ExtractSQLExpress(Session session)
      {
         try
         {
            // Extraction path is like C:\Users\user\AppData\Local\Temp\SQLEXPR_x64_ENU.exe
            String outputFile = session["TempFolder"] + session["SQLEXPRESSPACKAGE"];
            if (File.Exists(outputFile))
               File.Delete(outputFile);

            // Execute query on MSI database for the SQLExpressInstallerBinary
            View dbView = session.Database.OpenView("SELECT * FROM Binary WHERE Name = 'SQLExpressInstallerBinary'");
            dbView.Execute();

            // Fetch first record (assume one), get the "Data" field (index 2)
            Record rec = dbView.Fetch();
            Stream sqlBinary = rec.GetStream("Data");

            // If no data or data sizes don't match expected values...
            if (!(rec.GetDataSize("Data") > 0) || rec.GetDataSize("Data") != sqlBinary.Length)
               throw new Exception("Bad or no Data field found in record from msi's Binary table");

            // Copy data to outputfile
            using (FileStream fs = File.Create(outputFile))
            {
               sqlBinary.CopyTo(fs);
            }

            session.Log("SQLExpressInstallerBinary extracted to " + outputFile);
         }
         catch (Exception e)
         {
            session.Log("Exception: " + e.Message);
            return ActionResult.Failure;
         }
         return ActionResult.Success;
      }

      /// <summary>
      /// This will copy an extracted SQLExpress into the Database folder in BarTender, and cleanup files from %temp%
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success by default. Failure if unable to copy file into destination directory</returns>
      [CustomAction]
      public static ActionResult CopySQLExpress(Session session)
      {
         String tempFolder = session["TempFolder"];
         String databaseDir = session["SD_Database_Dir"];
         String sqlExprFile = "";

         try
         {
            sqlExprFile = Directory.GetFiles(tempFolder, "SQLEXPR_x??_ENU.exe", SearchOption.TopDirectoryOnly)[0];
         }
         catch { }

         try
         {
            // Copy the SQLExpress installer into the Database directory
            if (!String.IsNullOrEmpty(databaseDir) &&
                 !String.IsNullOrEmpty(tempFolder) &&
                 !String.IsNullOrEmpty(sqlExprFile) &&
                 File.Exists(sqlExprFile) &&
                 Directory.Exists(databaseDir))
            {
               try
               {
                  session.Log("Copying: " + sqlExprFile + " to " + Path.Combine(databaseDir, Path.GetFileName(sqlExprFile)));
                  File.Copy(sqlExprFile, Path.Combine(databaseDir, Path.GetFileName(sqlExprFile)), true);
               }
               catch (Exception e)
               {
                  session.Log("Copy failed: " + e.Message);
                  return ActionResult.Failure;
               }
            }
            try
            {
               // Clear RO flag if present, and cleanup the SQLEXPRESS installer from temp folder
               System.IO.FileAttributes attributes = File.GetAttributes(sqlExprFile);
               if ((attributes & System.IO.FileAttributes.ReadOnly) == System.IO.FileAttributes.ReadOnly)
               {
                  attributes = attributes & ~System.IO.FileAttributes.ReadOnly;
                  File.SetAttributes(sqlExprFile, attributes);
               }
               File.Delete(sqlExprFile);
            } catch (Exception e) { session.Log("Delete failed: " + e.Message); }

            try
            {
               // If the SQLExpress installer was run, let's cleanup any unpacked directory it may have made
               if (Directory.Exists(Path.Combine(tempFolder, Path.GetFileNameWithoutExtension(sqlExprFile))))
                  Directory.Delete(Path.Combine(tempFolder, Path.GetFileNameWithoutExtension(sqlExprFile)), true);
            } catch (Exception e) { session.Log("Delete Directory failed: " + e.Message); }
         }
         catch (Exception e)
         {
            session.Log("Unhandled exception: " + e.Message);
         }

         // Return Success
         return ActionResult.Success;
      }

      /// <summary>
      /// This will silently install a SQL Express 2014 instance with name "BarTender" if one
      /// does not already exist
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success if MSSQL instance is available and assured and ActionResult.Failure if there is a problem</returns>
      [CustomAction]
      public static ActionResult InstallSQLExpress(Session session)
      {
         // Return now if user controlled "INSTALLSQL" was not set
         if (session["INSTALLSQL"] != "true")
            return ActionResult.Success;

         try
         {
            // Check if MSSQL instance exists locally
            CustomActions.GetSQLInstanceState(session);

            // Install MSSQL instance if not detected
            if (session["SQLInstalled"] == "false")
            {
               System.Diagnostics.Process process = new System.Diagnostics.Process();
               process.StartInfo.UseShellExecute = true;
               process.StartInfo.Verb = "Open";
               process.StartInfo.RedirectStandardOutput = false;
               process.StartInfo.RedirectStandardError = false;
               process.StartInfo.CreateNoWindow = true;
               process.StartInfo.WorkingDirectory = session["TempFolder"];
               process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

               // Use first MSSQL Express installer found in the install path
               process.StartInfo.FileName = Directory.GetFiles(session["TempFolder"], "SQLEXPR_x??_ENU.exe", SearchOption.TopDirectoryOnly)[0];
               process.StartInfo.Arguments = @"/q /ACTION=Install /FEATURES=SQLEngine,FullText " +
                  @"/INSTANCENAME=" + defaultSQLInstanceName + @" /SQLSYSADMINACCOUNTS=""Builtin\Administrators"" ""NT AUTHORITY\SYSTEM"" " +
                  @"/SQLSVCACCOUNT=""NT AUTHORITY\SYSTEM"" /ADDCURRENTUSERASSQLADMIN /TCPENABLED=1 /IACCEPTSQLSERVERLICENSETERMS /HIDECONSOLE " +
                  @"/SkipInstallerRunCheck /UpdateEnabled=0 /SKIPRULES=RebootRequiredCheck SetupCompatibilityCheck NoRebootPackage";

               // Execute and wait
               session.Log("Installing MSSQL with commandline: \"" + process.StartInfo.FileName + "\" " + process.StartInfo.Arguments);
               process.Start();
               process.WaitForExit();

               // Log any errors
               if (process.ExitCode != 0)
               {
                  session.Log("Error during install, exit code was " + process.ExitCode);

                  // Attempt to find and fetch contents of the MSSQL installer's summary log
                  string pathToSummaryLog = "";
                  if (System.IO.File.Exists(@"C:\Program Files\Microsoft SQL Server\120\Setup Bootstrap\Log\Summary.txt"))
                     pathToSummaryLog = @"C:\Program Files\Microsoft SQL Server\120\Setup Bootstrap\Log\Summary.txt";
                  if (System.IO.File.Exists(@"C:\Program Files (x86)\Microsoft SQL Server\120\Setup Bootstrap\Log\Summary.txt"))
                     pathToSummaryLog = @"C:\Program Files (x86)\Microsoft SQL Server\120\Setup Bootstrap\Log\Summary.txt";
                  if (System.IO.File.Exists(@"C:\Program Files\Microsoft SQL Server\150\Setup Bootstrap\Log\Summary.txt"))
                     pathToSummaryLog = @"C:\Program Files\Microsoft SQL Server\150\Setup Bootstrap\Log\Summary.txt";
                  if (System.IO.File.Exists(@"C:\Program Files (x86)\Microsoft SQL Server\150\Setup Bootstrap\Log\Summary.txt"))
                     pathToSummaryLog = @"C:\Program Files (x86)\Microsoft SQL Server\150\Setup Bootstrap\Log\Summary.txt";
                  if (pathToSummaryLog.Length > 0)
                  {
                     session.Log(System.IO.File.ReadAllText(pathToSummaryLog));
                  }
                  else
                  {
                     session.Log("Failed to find a MSSQL installer log file");
                  }

                  // Exit now for all codes other than 3010 (which means "Passed but reboot required, see logs for details")
                  if (process.ExitCode != 3010)
                  {
                     session["SQLINSTALLFAILURE"] = "true";
                     return ActionResult.Failure;
                  }
               }
            }
            else
            {
               // SQL instance exists already
               session.Log("An MSSQL Server instance (local)\\" + defaultSQLInstanceName + " already exists, skipping installation");
            }

            // Sanity check the MSSQL service is running now
            ServiceController svc;
            try
            {
               svc = new ServiceController("MSSQL$" + defaultSQLInstanceName);
               svc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(60));
            }
            catch (System.ArgumentException e)
            {
               session.Log("Expected MSSQL service \"MSSQL$" + defaultSQLInstanceName + "\" does not exist: " + e.Message);
               session["SQLINSTALLFAILURE"] = "true";
               return ActionResult.Failure;
            }
            catch (System.ServiceProcess.TimeoutException)
            {
               session.Log("Timeout waiting for \"MSSQL$" + defaultSQLInstanceName + "\" status");
            }

            // Return success
            session["SQLInstalled"] = "true";
            return ActionResult.Success;
         }
         catch (Exception e)
         {
            session.Log("Unhandled exception: " + e.Message);
            session["SQLINSTALLFAILURE"] = "true";
            return ActionResult.Failure;
         }
      }

      /// <summary>
      /// This will create needed ACL permissions so DataBuilder can write to AppData folder
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult Success or Failure</returns>
      [CustomAction]
      public static ActionResult SetDataBuilderACL(Session session)
      {
         // ** DISABLE FUNCTIONALITY FOR NOW - Remove fully when DataBuilder is confirmed to be using LocalDB exclusively and no longer needs a full MSSQL instance **
         return ActionResult.Success;

         /*
         try
         {
            // Grant full permissions to MSSQL$INSTANCENAME to CommonAppData\Seagull folder
            //  e.g. cmd /C ""C:\Windows\System32\icacls.exe" "C:\ProgramData\Seagull" /grant "NT Service\MSSQL$Bartender":(OI)(CI)(F)"
            System.Diagnostics.Process icacls_process = new System.Diagnostics.Process();
            icacls_process.StartInfo.UseShellExecute = true;
            icacls_process.StartInfo.Verb = "Open";
            icacls_process.StartInfo.RedirectStandardOutput = false;
            icacls_process.StartInfo.RedirectStandardError = false;
            icacls_process.StartInfo.CreateNoWindow = true;
            icacls_process.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            icacls_process.StartInfo.FileName = "cmd.exe";
            icacls_process.StartInfo.Arguments = @"/c """"" + session["SystemFolder"] + @"\icacls.exe"" """ +
               session["CommonAppDataFolder"] + @"Seagull"" /grant ""NT Service\MSSQL$" + defaultSQLInstanceName + @""":(OI)(CI)(F)""";

            // Execute and wait
            session.Log("Granting MSSQL permissions to read/write to Seagull AppData folder with commandline: \"" + icacls_process.StartInfo.FileName + "\" " + icacls_process.StartInfo.Arguments);
            icacls_process.Start();
            icacls_process.WaitForExit();

            // Log any errors
            if (icacls_process.ExitCode != 0)
            {
               session.Log("Error during icalcs operation, exit code was " + icacls_process.ExitCode);
               return ActionResult.Failure;
            }

            // Return success
            return ActionResult.Success;
         }
         catch (Exception e)
         {
            session.Log("Unhandled exception: " + e.Message);
            return ActionResult.Failure;
         }
         */
      }

      /// <summary>
      /// This will create a default BarTender System Database, if a BSS DB does not exists.
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success regardless so that installer continues</returns>
      [CustomAction]
      public static ActionResult CreateDefaultSystemDB(Session session)
      {
         // Only create a default system database if there is not one configured already
         try
         {
            // Check SQL instance state and if SystemDB already configured
            CustomActions.GetSQLInstanceState(session);
            CustomActions.GetSystemDBState(session);

            // Make sure BSS is running
            if (!StartService("BarTender System Service"))
               session.Log("Unexpected problem starting BSS");

            // If SystemDbIsSetup is setup, return now
            if (session["SystemDbIsSetup"] == "true")
            {
               session.Log("BarTender System Database is already installed");
               return ActionResult.Success;
            }
            else if (session["SQLInstalled"] == "false")
            {
               session.Log("No local Microsoft SQL Server instance present, database will not be created automatically");
               return ActionResult.Success;
            }
            else
            {
               // Create BarTender System Database on default BarTender instance
               System.Diagnostics.Process process = new System.Diagnostics.Process();
               process.StartInfo.UseShellExecute = false;
               process.StartInfo.Verb = "runas";
               process.StartInfo.RedirectStandardOutput = true;
               process.StartInfo.RedirectStandardError = true;
               process.StartInfo.CreateNoWindow = true;
               process.StartInfo.WorkingDirectory = session["APPDIR"];
               process.StartInfo.FileName = "cmd.exe";

               process.StartInfo.Arguments = @"/C """"" + process.StartInfo.WorkingDirectory + @"\SystemDatabaseWizard.exe"" " +
               @"/Log=C:\ProgramData\Seagull\Installer\SystemDBLog.txt /Server=(local)\BarTender " +
               @"/Database=Datastore /type=Connect /BtSystem=127.0.0.1 /Silent""";

               // Execute and wait
               session.Log("Creating default BarTender System Database with command: " + process.StartInfo.Arguments);
               process.Start();
               process.WaitForExit();

               // Handle errors
               if (process.ExitCode != 1)
                  session.Log("ExitCode: " + process.ExitCode);
            }
         }
         catch (Exception e)
         {
            session.Log("Unhandled exception: " + e.Message);
         }
         return ActionResult.Success;
      }


      /// <summary>
      /// This will create named LocalDB instances if they do not exist on the local machine already.
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success or Failure</returns>
      [CustomAction]
      public static ActionResult ConfigureLocalDB(Session session)
      {
         try
         {
            ISqlLocalDbApi localDB = new SqlLocalDbApiWrapper();

            // Check that at least one LocalDB is installed
            if (!localDB.IsLocalDBInstalled())
            {
               session.Log("No LocalDB installed on system");
               return ActionResult.Failure;
            }

            // Display info on installed LocalDB versions and instances
            foreach (string version in localDB.Versions)
            {
               session.Log("LocalDB v" + version + " is installed");
            }
            foreach (string localDBInstanceName in localDB.GetInstanceNames())
            {
               session.Log("Found Instance: (localdb)\\" + localDBInstanceName);
            }

            // Create each of our instances
            foreach (KeyValuePair<String, String> kvp in CustomActions.localDBInstances)
            {
               // Check if this instance exists
               ISqlLocalDbInstanceInfo instance = localDB.GetInstanceInfo(kvp.Key);
               if (instance.Exists)
               {
                  if (instance.LocalDbVersion.Major + "." + instance.LocalDbVersion.Minor != kvp.Value)
                     session.Log("Error: " + instance.Name + " is running with version " + instance.LocalDbVersion.ToString());
                  else
                     session.Log("Instance " + instance.Name + " is already setup");
               }
               else
               {
                  session.Log("Creating instance: (localdb)\\" + instance.Name);
                  localDB.CreateInstance(kvp.Key, kvp.Value);
                  //localDB.StartInstance(kvp.Key);
               }
            }
            return ActionResult.Success;
         }
         catch (Exception e)
         {
            session.Log("Exception: " + e.ToString());
            return ActionResult.Failure;
         }
      }

      /// <summary>
      /// This will open the MSI database and extract the DotNetCoreHostingBundle binary from the binary table
      /// into the \AppData\Roaming\Seagull\Installer\prerequisites folder
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success if extraction was completed. ActionResult.Failure if there is a problem</returns>
      [CustomAction]
      public static ActionResult ExtractDotNetHostingBundle(Session session)
      {
         try
         {
            // Extraction path is like C:\Users\user\AppData\Local\Temp\dotnet-hosting-3.1.13-win.exe
            String outputFile = session["TempFolder"] + session["DOTNETHOSTINGPACKAGE"];
            if (File.Exists(outputFile))
               File.Delete(outputFile);

            // Execute query on MSI database for the DOTNETHostingInstallerBinary
            View dbView = session.Database.OpenView("SELECT * FROM Binary WHERE Name = 'DOTNETHostingInstallerBinary'");
            dbView.Execute();

            // Fetch first record (assume one), get the "Data" field (index 2)
            Record rec = dbView.Fetch();
            Stream dotNetBinary = rec.GetStream("Data");

            // If no data or data sizes don't match expected values...
            if (!(rec.GetDataSize("Data") > 0) || rec.GetDataSize("Data") != dotNetBinary.Length)
               throw new Exception("Bad or no Data field found in record from msi's Binary table");

            // Copy data to outputfile
            using (FileStream fs = File.Create(outputFile))
            {
               dotNetBinary.CopyTo(fs);
            }

            session.Log("DOTNETHostingInstallerBinary extracted to " + outputFile);
         }
         catch (Exception e)
         {
            session.Log("Exception: " + e.Message);
            return ActionResult.Failure;
         }
         return ActionResult.Success;
      }

      /// <summary>
      /// This will copy an extracted DotNetCoreHostingBundle into the prerequisites folder in \AppData\Roaming\Seagull\Installer\, and cleanup files from %temp%
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success by default. Failure if unable to copy file into destination directory</returns>
      [CustomAction]
      public static ActionResult CopyDotNetHostingBundle(Session session)
      {
         String tempFolder = session["TempFolder"];
         String prerequisitesDir = session["prerequisites_Dir"];
         String dotNetHostingFile = "";

         try
         {
            dotNetHostingFile = Directory.GetFiles(tempFolder, session["DOTNETHOSTINGPACKAGE"], SearchOption.TopDirectoryOnly)[0];
         }
         catch { }

         try
         {
            // Check if prerequisites folder exists for copying binary over from the TEMP folder
            if (!Directory.Exists(prerequisitesDir))
            {
               try
               {
                  Directory.CreateDirectory(prerequisitesDir);
               }
               catch (Exception ex)
               {
                  session.Log(ex.Message);
               }
            }
            // Copy the DotNetCoreHostingBundle installer into the prerequisites directory
            if (!String.IsNullOrEmpty(prerequisitesDir) &&
                 !String.IsNullOrEmpty(tempFolder) &&
                 !String.IsNullOrEmpty(dotNetHostingFile) &&
                 File.Exists(dotNetHostingFile) &&
                 Directory.Exists(prerequisitesDir))
            {
               try
               {
                  session.Log("Copying: " + dotNetHostingFile + " to " + Path.Combine(prerequisitesDir, Path.GetFileName(dotNetHostingFile)));
                  File.Copy(dotNetHostingFile, Path.Combine(prerequisitesDir, Path.GetFileName(dotNetHostingFile)), true);
               }
               catch (Exception e)
               {
                  session.Log("Copy failed: " + e.Message);
                  return ActionResult.Failure;
               }
            }
            try
            {
               // Clear RO flag if present, and cleanup the DotNetCoreHostingBundle installer from temp folder
               System.IO.FileAttributes attributes = File.GetAttributes(dotNetHostingFile);
               if ((attributes & System.IO.FileAttributes.ReadOnly) == System.IO.FileAttributes.ReadOnly)
               {
                  attributes = attributes & ~System.IO.FileAttributes.ReadOnly;
                  File.SetAttributes(dotNetHostingFile, attributes);
               }
               File.Delete(dotNetHostingFile);
            }
            catch (Exception e) { session.Log("Delete failed: " + e.Message); }

            try
            {
               // If the DotNetCoreHostingBundle installer was run, let's cleanup any unpacked directory it may have made
               if (Directory.Exists(Path.Combine(tempFolder, Path.GetFileNameWithoutExtension(dotNetHostingFile))))
                  Directory.Delete(Path.Combine(tempFolder, Path.GetFileNameWithoutExtension(dotNetHostingFile)), true);
            }
            catch (Exception e) { session.Log("Delete Directory failed: " + e.Message); }
         }
         catch (Exception e)
         {
            session.Log("Unhandled exception: " + e.Message);
         }

         // Return Success
         return ActionResult.Success;
      }

      /// <summary>
      /// This will install DotNetCore Hosting bundle prerequisite when we install BarTedner Print Portal (BPP). It will detect also if DotNetCore
      /// has already been installed and will force a repair once BPP is installed to enable the IIS portion and additional prepreqs 
      /// from DotNetCore
      /// </summary>
      /// <param name="session"></param>
      /// <returns>ActionResult.Success or Failure</returns>
      [CustomAction]
      public static ActionResult InstallDotNetCore(Session session)
      {
         // Adjust these two values to match version of the DotNet Core Hosting bundle and filename
         Version dotNetVersion = new Version("6.0.26.23605");
         string dotNetBinaryFilename = "dotnet-hosting-6.0.26-win.exe";

         Version sRegVersion = new Version("0.0.0.0");
         try
         {
            RegistryKey rk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            string registryPath = String.Format(@"SOFTWARE\Microsoft\ASP.NET Core\Shared Framework\v{0}.{1}\{0}.{1}.{2}", dotNetVersion.Major, dotNetVersion.Minor, dotNetVersion.Build);
            using (RegistryKey key = rk.OpenSubKey(registryPath))
            {
               if (key != null)
                  sRegVersion = new Version(key.GetValue("Version", String.Empty).ToString());
            }
            rk.Dispose();
         }
         catch (Exception e)
         {
            session.Log("Error obtaining DotNetHostingVersion from registry: " + e.Message);
            sRegVersion = new Version("0.0.0.0");
         }

         // Begin preparing commandline
         System.Diagnostics.Process prc = new System.Diagnostics.Process();
         prc.StartInfo.UseShellExecute = true;
         prc.StartInfo.Verb = "Open";
         prc.StartInfo.RedirectStandardOutput = false;
         prc.StartInfo.RedirectStandardError = false;
         prc.StartInfo.CreateNoWindow = true;
         prc.StartInfo.WorkingDirectory = session["prerequisites_Dir"];
         prc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
         prc.StartInfo.FileName = Directory.GetFiles(session["prerequisites_Dir"], dotNetBinaryFilename, SearchOption.TopDirectoryOnly)[0];
         try
         {
            if (sRegVersion < dotNetVersion)//run /quiet /norestart
               prc.StartInfo.Arguments = @" /quiet /norestart";
            else//run /quiet /repair /norestart
               prc.StartInfo.Arguments = @" /quiet /repair /norestart";
            // Finish arguments and execute
            prc.StartInfo.Arguments += "\"";
            session.Log("Running: " + prc.StartInfo.FileName + prc.StartInfo.Arguments);
            session.Log("sRegVersion: " + sRegVersion.ToString());
            prc.Start();
            prc.WaitForExit();

            if (prc.ExitCode != 3010)
            {
               return ActionResult.Failure;
            }
         }
         catch (Exception ex)
         {
            session.Log("Exception running dotnet-hosting.exe: " + ex.Message);
            return ActionResult.Failure;
         }
         return ActionResult.Success;
      }

      /// <summary>
      /// Migrates UniversalData from MS SQL Compact database to new upgrade LocalDB.
      /// ExportSqlCe40.exe is obtained from https://github.com/ErikEJ/SqlCeToolbox created by Erik Ejlskov Jensen (Microsoft MVP)
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult MigrateSqlCeDatabases(Session session)
      {
         String universalDataBaseName = "UniversalData_";
         String commonApplicationDataPath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

         try
         {
            string[] sqlCE40_dbFilesArray = Directory.GetFiles(commonApplicationDataPath, "Seagull\\Printer Maestro\\" + universalDataBaseName + "*.sdf");
            foreach (var sqlCE40_databaseFile in sqlCE40_dbFilesArray)
            {
               universalDataBaseName = Path.GetFileNameWithoutExtension(sqlCE40_databaseFile);
               String localDB_databaseFile = Path.Combine(commonApplicationDataPath, "Seagull\\Printer Maestro\\" + universalDataBaseName + ".mdf");
               String sqlDump_databaseFile = Path.Combine(commonApplicationDataPath, "Seagull\\Printer Maestro\\" + universalDataBaseName + ".sql");
               String assemblyFilePath = Path.Combine(session["APPDIR"], "Seagull.Data.Databuilder.dll");
               String sqlCE40_conn = string.Format(@"Data Source='{0}'; LCID=1033; Password='{1}'; Encrypt=TRUE;",
                  sqlCE40_databaseFile,
                  "{CD9A3EC8-EF01-4FED-B68A-6E1B564EEB27}");
               String localDB_conn = string.Format(@"Data Source=(LocalDB)\{0};Initial Catalog={1};AttachDbFilename={2};Integrated Security=True",
                  "MaestroServiceDB120",
                  universalDataBaseName,
                  localDB_databaseFile);

               // Return now if localDB file exists, as we should not overwrite it
               if (File.Exists(localDB_databaseFile))
               {
                  session.Log("A LocalDB UniversalData database file exists, skipping conversion.");
                  return ActionResult.Success;
               }

               // Return now if there is no SqlCompactCE database file in need of conversion
               if (!File.Exists(sqlCE40_databaseFile))
               {
                  session.Log("No MSSQL Compact UniversalData database file was found, skipping conversion.");
                  return ActionResult.Success;
               }

               // Return error if we can't find the Seagull.Data.Databuilder assembly
               if (!File.Exists(assemblyFilePath))
               {
                  session.Log("Error: Unable to find Seagull.Data.Databuilder assembly in installation path.");
                  return ActionResult.Failure;
               }

               // Create the localdb instance
               try
               {
                  ISqlLocalDbApi localDB = new SqlLocalDbApiWrapper();

                  // Check that at least one LocalDB is installed
                  if (!localDB.IsLocalDBInstalled())
                  {
                     session.Log("No LocalDB installed on system");
                     return ActionResult.Failure;
                  }

                  // Check if this instance exists
                  ISqlLocalDbInstanceInfo instance = localDB.GetInstanceInfo("MaestroServiceDB120");
                  if (instance.Exists)
                  {
                     session.Log("Instance " + instance.Name + " is already setup");
                  }
                  else
                  {
                     session.Log("Creating instance: (localdb)\\" + instance.Name);
                     localDB.CreateInstance("MaestroServiceDB120", "12.0");
                  }
               }
               catch (Exception ex)
               {
                  session.Log("Exception: " + ex.ToString());
                  return ActionResult.Failure;
               }

               // Use the ExportSqlCe40 tool to generate a sql dump of the existing UniversalData SqlCompactCE database
               System.Diagnostics.Process prc = new System.Diagnostics.Process();
               prc.StartInfo.UseShellExecute = true;
               prc.StartInfo.Verb = "Open";
               prc.StartInfo.RedirectStandardOutput = false;
               prc.StartInfo.RedirectStandardError = false;
               prc.StartInfo.CreateNoWindow = true;
               prc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
               prc.StartInfo.FileName = "ExportSqlCe40.exe";
               try
               {
                  prc.StartInfo.Arguments += " \"" + sqlCE40_conn + "\" \"" + sqlDump_databaseFile + "\"";
                  session.Log("Running: " + prc.StartInfo.FileName + " on " + universalDataBaseName + ".sdf");
                  prc.Start();
                  prc.WaitForExit();
               }
               catch (Exception ex)
               {
                  session.Log("Exception while running ExportSqlCe40.exe: " + ex.Message);
                  return ActionResult.Failure;
               }

               // Initialize the LocalDB database file for UniversalData
               try
               {
                  // LoadFile Assembly
                  Assembly assemblyDataBuilder = Assembly.LoadFile(assemblyFilePath);
                  Type sqlLocalServer = assemblyDataBuilder.GetType("Seagull.Data.DataBuilder.SqlServer.SqlLocalServer");
                  session.Log("Seagull.Data.Databuilder loaded from \"" + assemblyFilePath + "\"");
                  // SqlLocalServer.RawServerName = "MaestroServiceDB120";
                  sqlLocalServer.InvokeMember("set_RawServerName", BindingFlags.InvokeMethod, null, null, new object[] { "MaestroServiceDB120" });
                  // SqlLocalServer.CreateDatabaseOnFile( "UniversalData_7170", "C:\\ProgramData\\Seagull\\Printer Maestro\\UniversalData_7170.mdf", null, false );
                  MethodInfo methodCreateDatabaseOnFile = sqlLocalServer.GetMethod("CreateDatabaseOnFile");
                  methodCreateDatabaseOnFile.Invoke(null, new object[] { universalDataBaseName, localDB_databaseFile, null, false });
                  session.Log("Created database on file UniversalData_7170.mdf");
               }
               catch (Exception ex)
               {
                  session.Log("Assembly exception while creating database: " + ex.Message + "\r\nException: " + ex.ToString());
                  return ActionResult.Failure;
               }

               // Connect to LocalDB, and execute the sql dump commands to build out the schema and data.
               SqlConnection _dbConnection = new SqlConnection(localDB_conn);
               using (SqlCommand cmd = _dbConnection.CreateCommand() as SqlCommand)
               {
                  string script = File.ReadAllText(sqlDump_databaseFile);
                  try
                  {
                     _dbConnection.Open();
                     cmd.CommandText = script.Replace("GO", "");
                     int result = cmd.ExecuteNonQuery();
                     _dbConnection.Close();
                     session.Log("Migrated schema and data into UniversalData_7170.mdf");
                  }
                  catch (Exception ex)
                  {
                     session.Log("Exception while executing sql file: " + ex.Message + "\r\nException: " + ex.ToString());
                     return ActionResult.Failure;
                  }
               }

               // Remove the sql dump file and rename the sdf file so we do not attempt to process this action again.
               try
               {
                  File.Delete(sqlDump_databaseFile);
                  File.Move(sqlCE40_databaseFile, Path.ChangeExtension(sqlCE40_databaseFile, ".sd_"));
               }
               catch (Exception ex)
               {
                  session.Log("Exception while tidying Print Maestro folder: " + ex.Message + "\r\nException: " + ex.ToString());
               }
            }
         }
         catch (Exception e)
         {
            session.Log("Unhandled exception: " + e.Message);
         }
         return ActionResult.Success;
      }
   }
}
