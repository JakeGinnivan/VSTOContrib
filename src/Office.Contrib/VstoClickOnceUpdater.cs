using System;
using System.Deployment.Application;
using System.Diagnostics;
using System.Security;
using System.Security.Permissions;
using System.Security.Policy;
using System.IO;
using System.Text;
using Microsoft.Win32;
using Office.Contrib.Extensions;

namespace Office.Contrib
{
    /// <summary>
    /// Helper class for updating a VSTO add-in on demand. 
    /// Handles security issues. 
    /// See http://blogs.msdn.com/krimakey/archive/2008/04/18/click-once-forced-updates-in-vsto-ii-a-fuller-solution.aspx
    /// for more information on the issue this class solves
    /// Supports VSTO v3 and v4
    /// </summary>
    public class VstoClickOnceUpdater : ClickOnceUpdater
    {
        /// <summary>
        /// Updates the add-in.
        /// </summary>
        /// <param name="currentDeployment">The current application deployment.</param>
        /// <returns></returns>
        protected override UpdateResult UpdateApplication(ApplicationDeployment currentDeployment)
        {
            FixTrust(currentDeployment);
            return base.UpdateApplication(currentDeployment);
        }

        /// <summary>
        /// Updates the current deployment.
        /// </summary>
        /// <param name="deployment">The deployment.</param>
        /// <param name="message">The message.</param>
        /// <returns></returns>
        protected override bool UpdateCurrentDeployment(ApplicationDeployment deployment, ref string message)
        {
            //Call VSTOInstaller Explicitly in "Silent Mode"
            var installerPath = GetInstallerPath();
            if (installerPath == null)
            {
                message = "Cannot resolve VSTO Installer installation path";
                return false;                
            }
            var installerArgs = string.Format(" /S /I {0}", deployment.UpdateLocation.AbsoluteUri);

            var vstoInstallerOutput = new StringBuilder();

            var vstoStartInfo = new ProcessStartInfo(installerPath, installerArgs);
            var returnCode = vstoStartInfo.StartProcess((sender, e) => vstoInstallerOutput.Append((string) e.Data));

            message = vstoInstallerOutput.ToString();
            return returnCode == 0;
        }

        private static string GetInstallerPath()
        {
            var installerPath = string.Empty;
            //VSTO v4 (VS2010) now gives us a registry key, which is preferred.
            // if registry
            if (!GetInstallerPathFromRegistry(ref installerPath))
            {
                if (Directory.Exists(@"C:\Program Files (x86)"))
                {
                    if (Directory.Exists(@"C:\Program Files (x86)\Common Files\microsoft shared\VSTO\9.0"))
                        installerPath = @"C:\Program Files (x86)\Common Files\microsoft shared\VSTO\9.0\VSTOInstaller.exe";
                    else if (Directory.Exists(@"C:\Program Files (x86)\Common Files\microsoft shared\VSTO\8.0"))
                        installerPath = @"C:\Program Files (x86)\Common Files\microsoft shared\VSTO\8.0\VSTOInstaller.exe";
                    else
                        return null;
                }
                else
                {
                    if (Directory.Exists(@"C:\Program Files\Common Files\microsoft shared\VSTO\9.0"))
                        installerPath = @"C:\Program Files\Common Files\microsoft shared\VSTO\9.0\VSTOInstaller.exe";
                    else if (Directory.Exists(@"C:\Program Files\Common Files\microsoft shared\VSTO\8.0"))
                        installerPath = @"C:\Program Files\Common Files\microsoft shared\VSTO\8.0\VSTOInstaller.exe";
                    else
                        return null;
                }
            }

            return installerPath;
        }

        /// <summary>
        /// Gets the installer path from registry.
        /// This is the preferred method for VSTO 4.0
        /// </summary>
        /// <param name="installerPath">The installer path.</param>
        /// <returns></returns>
        private static bool GetInstallerPathFromRegistry(ref string installerPath)
        {
            var software = Registry.LocalMachine.OpenSubKey("SOFTWARE");
            if (software == null)
                return false;
            var microsoftKey = software.OpenSubKey("Microsoft");
            if (microsoftKey == null)
                return false;
            var vstoRuntimeSetupKey = microsoftKey.OpenSubKey("VSTO Runtime Setup");
            if(vstoRuntimeSetupKey == null)
                return false;
            var vsto4Key = vstoRuntimeSetupKey.OpenSubKey("v4");
            if (vsto4Key == null)
                return false;

            var path = vsto4Key.GetValue("InstallerPath");
            if (path == null)
                return false;

            installerPath = path.ToString();
            return true;

            //[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\VSTO Runtime Setup\v4]
            //“InstallerPath”=”C:\\Program Files\\Common Files\\Microsoft Shared\\VSTO\\10.0\\VSTOInstaller.exe”
        }

        private static void FixTrust(ApplicationDeployment currentDeployment)
        {
            //Create the appropriate Trust settings so the Application can do 
            //Click-Once Related updating
            var deploymentFullName = currentDeployment.UpdatedApplicationFullName;
            var appId = new ApplicationIdentity(deploymentFullName);
            var everything = new PermissionSet(PermissionState.Unrestricted);

            var trust = new ApplicationTrust(appId)
                            {
                                DefaultGrantSet = new PolicyStatement(everything),
                                IsApplicationTrustedToRun = true,
                                Persist = true
                            };

            ApplicationSecurityManager.UserApplicationTrusts.Add(trust);
        }
    }
}
