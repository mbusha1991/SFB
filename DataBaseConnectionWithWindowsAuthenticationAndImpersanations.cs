// This sample demonstrates the use of the WindowsIdentity class to impersonate a user.
// IMPORTANT NOTES:
// This sample requests the user to enter a password on the console screen.
// Because the console window does not support methods allowing the password to be masked,
// it will be visible to anyone viewing the screen.
// On Windows Vista and later this sample must be run as an administrator. 


using System;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Security.Permissions;
using Microsoft.Win32.SafeHandles;
using System.Runtime.ConstrainedExecution;
using System.Security;
using System.Data.SqlClient;

public class ImpersonationDemo
{
    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    public static extern bool LogonUser(String lpszUsername, String lpszDomain, String lpszPassword,
        int dwLogonType, int dwLogonProvider, out SafeTokenHandle phToken);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    public extern static bool CloseHandle(IntPtr handle);

    // Test harness.
    // If you incorporate this code into a DLL, be sure to demand FullTrust.
    [PermissionSetAttribute(SecurityAction.Demand, Name = "FullTrust")]
    public static SqlConnection impersonate(string uname, string password)
    {
        SafeTokenHandle safeTokenHandle;
        SqlConnection con= new SqlConnection();
        string[] username = uname.Split('\\');
        string domain = username[0];
        string name = username[1];
        try
        {
           
            const int LOGON32_PROVIDER_DEFAULT = 0;
            //This parameter causes LogonUser to create a primary token.
            const int LOGON32_LOGON_INTERACTIVE = 2;

            // Call LogonUser to obtain a handle to an access token.
            bool returnValue = LogonUser(name, domain, password,
                LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT,
                out safeTokenHandle);

            Console.WriteLine("LogonUser called.");

            if (false == returnValue)
            {
                int ret = Marshal.GetLastWin32Error();
                Console.WriteLine("LogonUser failed with error code : {0}", ret);
                throw new System.ComponentModel.Win32Exception(ret);
            }
            using (safeTokenHandle)
            {
                Console.WriteLine("Did LogonUser Succeed? " + (returnValue ? "Yes" : "No"));
                Console.WriteLine("Value of Windows NT token: " + safeTokenHandle);
                 
                // Use the token handle returned by LogonUser.
                using (WindowsIdentity newId = new WindowsIdentity(safeTokenHandle.DangerousGetHandle()))
                {
                    using (WindowsImpersonationContext impersonatedUser = newId.Impersonate())
                    {
                        con = new SqlConnection("Data Source=FE0PLYMD03.de.bosch.com;Initial Catalog=DE_LcsCDR;Persist Security Info=False;Integrated Security=SSPI");
                        
                            try
                            {
                                string cur_name = WindowsIdentity.GetCurrent().Name;
                                con.Open();                             
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Exception occurred. " + ex.Message);
                            }
                    }
                }
             }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Exception occurred. " + ex.Message);
        }
        return con;
    }
}
public sealed class SafeTokenHandle : SafeHandleZeroOrMinusOneIsInvalid
{
    private SafeTokenHandle()
        : base(true)
    {
    }

    [DllImport("kernel32.dll")]
    [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
    [SuppressUnmanagedCodeSecurity]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    protected override bool ReleaseHandle()
    {
        return CloseHandle(handle);
    }
}

