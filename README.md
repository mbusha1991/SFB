# SFB
//Pre requisites:
// To run powershell in the remote machine
//Open poweshell in admin mode and run this command in powershell
//set-item wsman:\localhost\Client\TrustedHosts -value *
/*
 * Author: Usha M Basavaiah
 * Date: 3/03/2018
 * Description: Powershell File Execution 
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.IO;
namespace SFBTesting.LibraryFunctions
{
    public class PowerShellExec
    {
        public bool RunPowerShellScript(string scriptFullpath, out string output, out string errors)
        {
            errors = string.Empty;
            try
            {
                var runspace = RunspaceFactory.CreateRunspace();
                return RunPowerShellScriptInternal(scriptFullpath, out output, out errors, runspace);
            }
            catch (Exception e)
            {
                errors += "Error occurred while Creating Runspace: " + e.Message + Environment.NewLine + e.Source + Environment.NewLine + e.StackTrace + Environment.NewLine + e.InnerException;
                output = string.Empty;
                return false;
                throw e;
            }
        }
        public bool RunPowerShellScriptRemote(string scriptFullpath, string computer, string username, string password, out string output, out string errors)
        {
            errors = string.Empty;
            try
            {
                var filename = Path.GetFileName(scriptFullpath);
                var credentials = new PSCredential(username, convertToSecureString(password));
                //port – Thanks to the blog post at http://blogs.msdn.com/b/wmi/archive/2009/07/22/new-default-ports-for-ws-management-and-powershell-remoting.aspx
                //we know what port numbers PowerShell remoting uses. 5985 if you are not using SSL, and 5986 if you are using SSL
                var connectionInfo = new WSManConnectionInfo(false, computer, 5985, "/wsman", "http://schemas.microsoft.com/powershell/Microsoft.PowerShell", credentials);
                var runspace = RunspaceFactory.CreateRunspace(connectionInfo);
                var remoteScriptFullpath = "\\\\01hw894947\\D$\\Temp\\" + this.GetType().Module.Name + "\\" + Guid.NewGuid() + "\\";
                remoteScriptFullpath = Path.GetFullPath(remoteScriptFullpath);
                if (!Directory.Exists(remoteScriptFullpath))
                {
                    Directory.CreateDirectory(remoteScriptFullpath);
                }
                remoteScriptFullpath += filename;
                File.Copy(scriptFullpath, remoteScriptFullpath, true);
                return RunPowerShellScriptInternal(remoteScriptFullpath, out output, out errors, runspace);
            }
            catch (Exception e)
            {
                errors += "Error occurred while Remote System Connection: " + e.Message + Environment.NewLine + e.Source + Environment.NewLine + e.StackTrace + Environment.NewLine + e.InnerException;
                output = string.Empty;
                return false;
                throw e;
            }
        }
        private System.Security.SecureString convertToSecureString(string password)
        {
            System.Security.SecureString secure = new System.Security.SecureString();
            try
            {
                if (!string.IsNullOrEmpty(password))
                {
                    foreach (char c in password)
                    {
                        secure.AppendChar(c);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\r\n" + ex.StackTrace + "\r\n" + ex.InnerException);
                throw ex;
            }
            return secure;
        }
        private bool RunPowerShellScriptInternal(string scriptFullpath, out string output, out string errors, Runspace runspace)
        {
            bool isExecutionSuccessful = true;
            output = string.Empty;
            errors = string.Empty;
            Collection<PSObject> results = new Collection<PSObject>();
            Pipeline pipeline = null;
            StringBuilder stringBuilder = new StringBuilder();
            try
            {
                runspace.Open();
                pipeline = runspace.CreatePipeline();
                var cmd = new Command(scriptFullpath);
                pipeline.Commands.Add(cmd);
                pipeline.Commands.Add("Out-String");
                results = pipeline.Invoke();
                foreach (PSObject obj in results)
                {
                    stringBuilder.AppendLine(obj.ToString());
                }
                if (pipeline.Error.Count > 0)
                {
                    errors += String.Join(Environment.NewLine, pipeline.Error.ReadToEnd().Select(e => e.ToString()));
                    isExecutionSuccessful = false;
                }
            }
            catch (Exception e)
            {
                errors += "Error occurred in PowerShell script: " + e.Message + Environment.NewLine + e.Source + Environment.NewLine + e.StackTrace + Environment.NewLine + e.InnerException;
                isExecutionSuccessful = false;
                throw e;
            }
            finally
            {
                output = stringBuilder.ToString();
                pipeline.Dispose();
                runspace.Dispose();
            }
            return isExecutionSuccessful;
        }
    }
}
