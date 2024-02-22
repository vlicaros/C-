using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;

namespace CustomActions
{
   public partial class CustomActions
   {
      /// <summary>
      /// Custom action to handle and show setup interrupted & user cancel dialogs.
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult SetupInterrupted(Session session)
      {
         // Set the msiexec process window handle
         wizard.WaitDialogProcessName = GetWindowHandleByProcessName("msiexec");

         // Attempt to hide the parent window
         ShowWindow(GetWindowHandleByProcessName("msiexec"), SW_HIDE);

         // Set install languge for activation wizard
         try
         {
            wizard.InstallLanguage = session["ProductLanguage"];
            session.Log("Product languge detected: " + session["ProductLanguage"]);
         }
         catch (Exception ex)
         {
            session.Log("Exception has occurred:");
            session.Log(ex.Message);
         }

         // Set setup inturrupted flag
         wizard.SetupInterrupted = true;

         // Set MSSQL install failure flag
         wizard.SQLInstallFailure = (session["SQLINSTALLFAILURE"] == "true");

         // Set RebootRequired flag
         wizard.RebootRequired = (session["WindowsUpdateRebootRequired"] == "true");

         // Show InstallWizard form
         if (session["UILevel"] != "2")
            wizard.Show();

         return ActionResult.Success;
      }

      /// <summary>
      /// Custom action to show the deactivation dialog during an Uninstall.
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult DeactivationUninstall(Session session)
      {
         // Find msiexec process window handle
         wizard.WaitDialogProcessName = GetWindowHandleByProcessName("msiexec");

         // Show deactivation form
         wizard.DoDeactivation = true;
         // Show InstallWizard form
         if (session["UILevel"] != "2")
            wizard.Show();

         // If the user choose to cancel install, raise error to the installer (SIGMA-1538)
         if (wizard.IsCancel)
         {
            return ActionResult.UserExit;
         }

         return ActionResult.Success;
      }

      /// <summary>
      /// Custom action to show post install options.
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult PostInstall(Session session)
      {
         // Set the msiexec process window handle
         wizard.WaitDialogProcessName = GetWindowHandleByProcessName("msiexec");

         // Set postinstall and load session data into wizard
         wizard.IsPostInstall = true;
         PrepareWizard(session);

         // REMOVE/UNINSTALL
         if (session["REMOVE"] == "ALL" || session["InstallMode"] == "Remove")
         {
            wizard.RemoveInstall = true;
            if (session["UILevel"] != "2")
               wizard.Show();
            return ActionResult.Success;
         }
         // MODIFY
         else if (session["IS_MODIFY"] == "TRUE")
         {
            wizard.ModifyInstall = true;
         }
         // REPAIR
         else if (session["InstallMode"] == "Repair")
         {
            wizard.RepairInstall = true;
         }
         // INSTALL
         else
         {
            // Spin up BSS
            if (!StartService("BarTender System Service"))
            {
               session.Log("BarTender System Service could not be started");
               wizard.ServiceError = true;
            }
         }

         // Show InstallWizard post install form
         if (session["UILevel"] != "2")
         {
            wizard.Show();
         }

         // Detect Spanish (Mexico) and default to Spanish International Sort
         String awProductLangID = session["ProductLanguage"];
         if (session["ProductLanguage"] == "3082")
             awProductLangID = "1034";
          
         // Attempt to activate using provided properties
         if (session["PKC"].Trim() != string.Empty)
         {
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
               prc.StartInfo.Arguments = @"/C """"" + System.IO.Path.Combine(session["INSTALLDIR"], "ActivationWizard.exe") + @"""" +
                  String.Format(" Installer {0} Activate PKC={1}", awProductLangID, session["PKC"]);

               // Add the LSIP and LSPORT if the BLS property was used
               if (session["BLS"].Trim() != string.Empty)
               {
                  string[] uri = session["BLS"].Split(':');
                  prc.StartInfo.Arguments += " LSIP=" + uri[0] + " LSPORT=" + uri[1];
               }

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
         else
         {
            // If we are not already activated and we're not running in silent install, open ActivationWizard
            if (session["UILevel"] != "2" && ! wizard.IsActivated)
            {
               try
               {
                  // Hide the msiexec window now
                  ShowWindow(GetWindowHandleByProcessName("msiexec"), SW_HIDE);

                  // Spin up AW ActivateNow
                  System.Diagnostics.Process prc = new System.Diagnostics.Process();
                  prc.StartInfo.UseShellExecute = false;
                  prc.StartInfo.FileName = "cmd.exe";
                  prc.StartInfo.Verb = "runas";
                  prc.StartInfo.RedirectStandardOutput = true;
                  prc.StartInfo.RedirectStandardError = true;
                  prc.StartInfo.CreateNoWindow = true;
                  prc.StartInfo.WorkingDirectory = session["INSTALLDIR"];
                  prc.StartInfo.Arguments = @"/C """"" + System.IO.Path.Combine(session["INSTALLDIR"], "ActivationWizard.exe") + @"""" +
                     String.Format(" Installer {0} ActivateNow", awProductLangID) + @"""";

                  prc.Start();
               }
               catch (Exception ex)
               {
                  session.Log("Exception running ActivationWizard: " + ex.Message);
               }
            }
         }

         // Always return success
         return ActionResult.Success;
      }

      /// <summary>
      /// Custom action to handle maintenance/repair/uninstall modes.
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult MaintenanceOptions(Session session)
      {
         // Set the msiexec process window handle
         wizard.WaitDialogProcessName = GetWindowHandleByProcessName("msiexec");

         // Attempt to hide the parent window
         ShowWindow(GetWindowHandleByProcessName("msiexec"), SW_HIDE);

         // Set maintenance flag and load session data
         // REMOVE/UNINSTALL
         if (session["REMOVE"] == "ALL" || session["InstallMode"] == "Remove")
         {
            wizard.RemoveInstall = true;
         }
         else
         {
             wizard.ModifyInstall = true;
             PrepareWizard(session);
             wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.SameVersion;
         }
         
         // Show InstallWizard maintenance form
         if (session["UILevel"] != "2")
         {
             // Find the features that are already installed
             List<InstallFeatures> currentFeatures = new List<InstallFeatures>();

             // Find the features that are already installed.
             session.Log("Current feature states:");
             foreach (FeatureInfo fi in session.Features)
             {
                InstallFeatures outValue;
                if (fi.CurrentState == InstallState.Local && Enum.TryParse<InstallFeatures>(fi.Name.ToString(), true, out outValue))
                {
                   currentFeatures.Add(outValue);
                   session.Log(" Feature {0} is currently InstallState.Local", outValue.ToString());
                }
             }

             // Figure out InstallType for the wizard
             if (currentFeatures.Contains(InstallFeatures.BPP))
                wizard.InstallType = Seagull.InstallWizard.InstallTypeSelection.PrintPortal;
             else if (currentFeatures.Contains(InstallFeatures.BarTender))
                wizard.InstallType = Seagull.InstallWizard.InstallTypeSelection.BarTender;
             else
                wizard.InstallType = Seagull.InstallWizard.InstallTypeSelection.LicensingService;

            wizard.Show();
         }
         else
         {
              // Set InstallType from commandline for the silent install
             if (session["FEATURE"].ToLowerInvariant() == "printportal")
                wizard.InstallType = Seagull.InstallWizard.InstallTypeSelection.PrintPortal;
             else if (session["FEATURE"].ToLowerInvariant() == "bartender")
                wizard.InstallType = Seagull.InstallWizard.InstallTypeSelection.BarTender;
             else
                wizard.InstallType = Seagull.InstallWizard.InstallTypeSelection.LicensingService;
         }

         // If the user choose to cancel install, raise error to the installer
         if (wizard.IsCancel)
         {
            session.Log("User has canceled");
            return ActionResult.UserExit;
         }

         // Set BACKUP_FORMATS and BACKUP_PREFERENCES property
         try
         {
            if (wizard.BackupFormats)
            {
               session["BACKUP_FORMATS"] = "1";
               session.Log("Setting backup_formats: 1");
            }

            if (wizard.BackupPreferences)
            {
               session["BACKUP_PREFERENCES"] = "1";
               session.Log("Setting backup_formats: 1");
            }
            else
            {
               session["BACKUP_PREFERENCES"] = "0";
               session.Log("Setting backup_formats: 0");
            }
         }
         catch (Exception ex)
         {
            session.Log("Exception has occurred:");
            session.Log(ex.Message);
            return ActionResult.Failure;
         }

         // Set PRINTPORTAL_ACCOUNT_PASSWORD property
         try
         {
             if (session["UILevel"] != "2")
             {
                session["PRINTPORTAL_ACCOUNT_PASSWORD"] = wizard.BPPPassword.ToString();
                session.Log("Setting PRINTPORTAL_ACCOUNT_PASSWORD to: " + wizard.BPPPassword.ToString());
             }
         }
         catch (Exception ex)
         {
            session.Log("Exception has occurred:");
            session.Log(ex.Message);
            return ActionResult.Failure;
         }

         // Set user selected SQL install option
         try
         {
            if (wizard.RepairInstall)
               session["INSTALLSQL"] = "false";
            else if (session["INSTALLSQL"].Trim() == string.Empty)
               session["INSTALLSQL"] = (wizard.AddSQLServer ? "true" : "false");
            session.Log("Installer will " + ((session["INSTALLSQL"] == "true") ? "" : "NOT ") + "attempt to install Microsoft SQL Express Server locally");
         }
         catch (Exception ex)
         {
            session.Log("Exception has occurred:");
            session.Log(ex.Message);
            return ActionResult.Failure;
         }

         // Handle Repair operations
         if (wizard.RepairInstall)
         {
            // Set maintenance modes
            session["InstallMode"] = "Repair";
            session["AI_INSTALL_MODE"] = "Repair";
            session["REINSTALLMODE"] = "ocmusv";
            session["REINSTALL"] = "ALL";

            // Set progress bar properies
            session["Progress1"] = (session["CtrlEvtRepairing"] != null) ? session["CtrlEvtRepairing"] : "Repairing";
            session["Progress2"] = (session["CtrlEvtrepairs"] != null) ? session["CtrlEvtrepairs"] : "repairs";
            session["ProgressType1"] = "Installing";
            session["ProgressType2"] = "installed";
            session["ProgressType3"] = "installs";

            return ActionResult.Success;
         }

         // Handle Remove operations
         if (wizard.RemoveInstall)
         {
            // Set maintenance modes
            session["InstallMode"] = "Remove";
            session["AI_INSTALL_MODE"] = "Remove";
            session["REMOVE"] = "ALL";
            session["EnableRollback"] = "False";

            // Set progress bar properties
            session["Progress1"] = (session["CtrlEvtRemoving"] != null) ? session["CtrlEvtRemoving"] : "Removing";
            session["Progress2"] = (session["CtrlEvtremoves"] != null) ? session["CtrlEvtremoves"] : "removes";
            session["ProgressType1"] = "Uninstalling";
            session["ProgressType2"] = "uninstalled";
            session["ProgressType3"] = "uninstalls";

            return ActionResult.Success;
         }

         // Handle modify operations
         session["IS_MODIFY"] = "TRUE";

         // Get the newly selected feature set
         List<InstallFeatures> desiredFeatures = GetFeatures(wizard.InstallType.ToString(), "");

         // Set the features
         foreach (InstallFeatures feature in Enum.GetValues(typeof(InstallFeatures)))
         {
            try
            {
               if (desiredFeatures.Contains(feature))
               {
                  session.Features[feature.ToString()].RequestState = InstallState.Local;
                  session.Log("Feature {0} has been marked InstallState.Local", feature.ToString());
               }
               else
               {
                  session.Features[feature.ToString()].RequestState = InstallState.Absent;
                  session.Log("Feature {0} has been marked InstallState.Absent", feature.ToString());
               }
            }
            catch (Exception ex)
            {
               session.Log("Exception has occurred: " + ex.Message);
            }
         }

         return ActionResult.Success;
      }
      /// <summary>
      /// Custom action to detect installed version.
      /// </summary>
      /// <param name="session"></param>
      /// <returns></returns>
      [CustomAction]
      public static ActionResult DetectInstallVersion(Session session)
      {
         //// Existing version info
         //wizard.InstalledVersion = session["BT_INSTALLED_VERSION_STRING"];
         //string currentInstalledVersion = session["BT_INSTALLED_VERSION"].Trim();
         //if (!string.IsNullOrEmpty(currentInstalledVersion))
         //{
         //   Version installerVersion = new Version();
         //   Version currentVersion = new Version();

         //   session.Log(string.Format("Currently installed version: {0}", currentInstalledVersion));
         //   session.Log(string.Format("Installer's version: {0}", wizard.FullVersion));
         //   try
         //   {
         //      installerVersion = new Version(wizard.FullVersion);
         //      currentVersion = new Version(currentInstalledVersion);
         //   }
         //   catch (Exception ex)
         //   {
         //      session.Log("Exception has occurred:");
         //      session.Log(ex.Message);
         //   }

         //   if (currentVersion != new Version() && installerVersion != new Version())
         //   {
         //      if (installerVersion == currentVersion)
         //      {
         //         // Versions are the same
         //         session.Log(string.Format("Installer versions match"));
         //      }
         //      else if (installerVersion > currentVersion)
         //      {
         //         if ((session["x86_or_x64"] == "x86" && Environment.Is64BitOperatingSystem) || (session["x86_or_x64"] == "x64" && !Environment.Is64BitOperatingSystem))
         //         {
         //            // Installer is a cross platform upgrade, set AI_UPGRADE to trigger UNINSTALL_PREVINSTALL_XPLATFORM_UPGRADE
         //            session.Log(string.Format("This installer is a cross-platform upgrade"));
         //            session["AI_UPGRADE"] = "Yes";
         //            wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.OlderVersion;
         //         }
         //         else
         //         {
         //            session.Log(string.Format("This installer is an upgrade"));
         //            session["AI_UPGRADE"] = "Yes";
         //            // Installer is newer, so tell wizard that the previous version is 'older'
         //            wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.OlderVersion;
         //         }
         //      }
         //      else if (installerVersion < currentVersion)
         //      {
         //         // Installer is older, so tell wizard that the previous version is 'newer'
         //         session.Log(string.Format("This installer is a downgrade"));
         //         session["AI_DOWNGRADE"] = "Yes";
         //         session["REINSTALLMODE"] = "amus";
         //         wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.NewerVersion;
         //      }
         //   }
         //   else
         //   {
         //      session["AI_UPGRADE"] = "Yes";
         //      // Installer is newer, so tell wizard that the previous version is 'older'
         //      wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.OlderVersion;
         //      session.Log("Unable to parse either the current installation's or installer's version numbers");
         //   }
         //}
         session["AI_UPGRADE"] = "Yes";
         session["MIGRATE"] = session["OLDPRODUCTS"];
         // Installer is newer, so tell wizard that the previous version is 'older'
         wizard.PreviousInstall = Seagull.InstallWizard.PreviousInstallType.OlderVersion;
         return ActionResult.Success;
      }
         /// <summary>
         /// Custom action to handle install options during normal (InstallWizard) or silent installs.
         /// </summary>
         /// <param name="session"></param>
         /// <returns></returns>
         [CustomAction]
      public static ActionResult InstallOptions(Session session)
      {
         // Set the msiexec process window handle
         wizard.WaitDialogProcessName = GetWindowHandleByProcessName("msiexec");

         // Attempt to hide the parent window
         ShowWindow(GetWindowHandleByProcessName("msiexec"), SW_HIDE);

         // Load session data into wizard
         PrepareWizard(session);

         // Display InstallWizard
         if (session["UILevel"] == "5")
         {
            session.Log("InstallWizard Started " + System.DateTime.Now.ToString());
            wizard.Show();
            session.Log("InstallWizard Finished " + System.DateTime.Now.ToString());
         }

         // If the user choose to cancel install, raise error to the installer
         if (wizard.IsCancel)
         {
            session.Log("User has canceled");
            return ActionResult.UserExit;
         }

         // Sanitize and convert install directories
         if (wizard.InstallPath.Trim().ToLowerInvariant() != session["INSTALLDIR"].Trim().ToLowerInvariant())
         {
            string sep = System.IO.Path.DirectorySeparatorChar.ToString();
            session["INSTALLDIR"] = wizard.InstallPath.Trim() + ((wizard.InstallPath.Trim().EndsWith(sep)) ? "" : sep);
         }
         session["TARGETDIR"] = session["INSTALLDIR"];
         session["APPDIR"] = session["INSTALLDIR"];

         // Set account user name & password
         if (!string.IsNullOrEmpty(wizard.AccountUserName))
         {
            session["ACCOUNT_NAME_FULL"] = wizard.AccountUserName;
            session.Log(string.Format("Setting Account Username: "), wizard.AccountUserName);
         }
         if (!string.IsNullOrEmpty(wizard.AccountUserName))
         {
            session["ACCOUNT_PASSWORD"] = wizard.AccountPassword;
            session.Log(string.Format("Setting Account Password: "), wizard.AccountPassword);
         }

         // Set if new user account needs to be created
         if (wizard.CreateAccount)
         {
            session["CREATE_ACCOUNT"] = "TRUE";
            session.Log("Setting Create Account Flag: TRUE");
         }

         // Set PRINTPORTAL_ACCOUNT_PASSWORD property
         if (wizard.BPPPassword != string.Empty && wizard.BPPPassword != session["PRINTPORTAL_ACCOUNT_PASSWORD"].Trim())
         {
            session["PRINTPORTAL_ACCOUNT_PASSWORD"] = wizard.BPPPassword.ToString();
            session.Log("Setting PRINTPORTAL_ACCOUNT_PASSWORD to: " + wizard.BPPPassword.ToString());
         }

         // Set user selected SQL install option
         if (session["INSTALLSQL"].Trim() == string.Empty)
            session["INSTALLSQL"] = (wizard.AddSQLServer ? "true" : "false");
         session.Log("Installer will " + ((session["INSTALLSQL"] == "true") ? "" : "NOT ") + "attempt to install Microsoft SQL Express Server locally");

         // Set IISROOTFOLDER property
         if (wizard.IISLocation != string.Empty)
         {
            session["IISROOTFOLDER"] = wizard.IISLocation;
            session.Log("Setting IISROOTFOLDER to: " + wizard.IISLocation);
         }

         // Set IISVERSION property
         if (wizard.IISVersion != string.Empty)
         {
            session["IIS_VERSION"] = wizard.IISVersion;
            session.Log("Setting IIS_VERSION to: " + wizard.IISVersion);
         }

         // Try to force an appsearch to refresh IIS properties to the installer
         try
         {
            session.Log("Executing appsearch");
            session.DoAction("AppSearch");
         }
         catch { }

         ////// SET FEATURES //////

         // Parse the feature list selected by the user either via UI or commandline
         List<InstallFeatures> featureList = GetFeatures(wizard.InstallType.ToString(), "");

         // Override and sanitize featureList using provided ADDLOCAL and REMOVE properties, if available
         if (session["ADDLOCAL"].Trim() != string.Empty || session["REMOVE"].Trim() != string.Empty)
            featureList = GetFeatures(session["ADDLOCAL"].ToString(), session["REMOVE"].ToString());

         // Make certain we have at least one feature to install
         if (featureList.Count == 0)
         {
            session.Log("No features are to be installed.  Please run installer in full UI mode or check the FEATURE and REMOVE property usage");
            return ActionResult.Failure;
         }

         // Set the features from the santized featureList
         foreach (InstallFeatures feature in Enum.GetValues(typeof(InstallFeatures)))
         {
            try
            {
               if (featureList.Contains(feature))
               {
                  session.Features[feature.ToString()].RequestState = InstallState.Local;
                  session.Log("Marking feature {0} as InstallState.Local", feature.ToString());
               }
               else
               {
                  session.Features[feature.ToString()].RequestState = InstallState.Absent;
                  session.Log("Marking feature {0} as InstallState.Absent", feature.ToString());
               }
            }
            catch (Exception ex)
            {
               session.Log("Exception has occurred: " + ex.Message);
            }
         }

        // Override for the ADDLOCAL or REMOVE feature properties
         if (session["ADDLOCAL"].Trim() != string.Empty || session["REMOVE"].Trim() != string.Empty)
         {
            // Apply ADDLOCAL (unsanitized entries)
            foreach (string addLocalFeature in session["ADDLOCAL"].Split(','))
         {
               try
               {
                  if (session.Features[addLocalFeature].RequestState != InstallState.Local)
                  {
                     session.Features[addLocalFeature].RequestState = InstallState.Local;
                     session.Log("Marking feature {0} as InstallState.Local", addLocalFeature.ToString());
                  }
               }
               catch (Exception ex)
               {
                  session.Log("Exception has occurred: " + ex.Message);
               }
            }

            // Apply REMOVE (unsanitized entries)
            foreach (string removeLocalFeature in session["REMOVE"].Split(','))
            {
               try
               {
                  if (session.Features[removeLocalFeature].RequestState != InstallState.Absent)
                  {
                     session.Features[removeLocalFeature].RequestState = InstallState.Absent;
                     session.Log("Marking feature {0} as InstallState.Absent", removeLocalFeature.ToString());
                  }
               }
               catch (Exception ex)
               {
                  session.Log("Exception has occurred: " + ex.Message);
               }
         }

            // Set the final ADDLOCAL/REMOVE properties
            List<string> addLocal = new List<string>();
            List<string> remove = new List<string>();
            foreach (FeatureInfo fi in session.Features)
            {
               if (fi.RequestState == InstallState.Local) // || fi.RequestState == InstallState.Default)
                  addLocal.Add(fi.Name.ToString());
               else if (fi.RequestState == InstallState.Absent) // || fi.RequestState == InstallState.Unknown)
                  remove.Add(fi.Name.ToString());
            }
            session.Log("Applying ADDLOCAL and REMOVE properties");
            session["ADDLOCAL"] = string.Join(",", addLocal);
            session["REMOVE"] = string.Join(",", remove);
            session["Preselected"] = "1";
         }

         // Display finalized feature states
         session.Log("Features:");
         foreach (FeatureInfo fi in session.Features)
         {
            session.Log("{0} is {1}", fi.Name.ToString(), fi.RequestState.ToString());
         }

         // Ensure PRINTPORTAL_ACCOUNT_PASSWORD is available if BPP was used
         if (session.Features["BPP"].RequestState == InstallState.Local && session["PRINTPORTAL_ACCOUNT_PASSWORD"].Trim() == string.Empty)
         {
            session.Features["BPP"].RequestState = InstallState.Absent;
            session.Log("Using the PrintPortal or BPP features require you to provide a PRINTPORTAL_ACCOUNT_PASSWORD property");
            return ActionResult.Failure;
         }

         // Return success to the installer
         return ActionResult.Success;
      }

   }
}