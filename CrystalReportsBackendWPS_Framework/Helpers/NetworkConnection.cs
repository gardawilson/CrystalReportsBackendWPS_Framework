using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;

namespace CrystalReportsBackendWPS_Framework.Helpers
{
    public class NetworkConnection : IDisposable
    {
        string _networkName;

        public NetworkConnection(string networkName, NetworkCredential credentials)
        {
            _networkName = networkName;

            // LOG
            string logPath = @"C:\Temp\network_connection_debug.txt";
            if (!Directory.Exists(@"C:\Temp")) Directory.CreateDirectory(@"C:\Temp");

            File.AppendAllText(logPath, $"[{DateTime.Now}] Disconnecting any previous connection to {_networkName}\n");

            // DISCONNECT previous connections to the same server
            string uncRoot = GetUncRoot(_networkName); // e.g. \\192.168.10.100

            Process.Start("cmd.exe", $"/C net use {uncRoot} /delete").WaitForExit();
            System.Threading.Thread.Sleep(1000); // wait a bit

            File.AppendAllText(logPath, $"[{DateTime.Now}] Connecting to {_networkName} as {credentials.UserName}\n");

            var netResource = new NetResource()
            {
                Scope = ResourceScope.GlobalNetwork,
                ResourceType = ResourceType.Disk,
                DisplayType = ResourceDisplaytype.Share,
                RemoteName = networkName
            };

            string userName = string.IsNullOrEmpty(credentials.Domain)
                ? credentials.UserName
                : $@"{credentials.Domain}\{credentials.UserName}";

            int result = WNetAddConnection2(netResource, credentials.Password, userName, 0);
            if (result != 0)
            {
                string errorMsg = new Win32Exception(result).Message;
                File.AppendAllText(logPath, $"[{DateTime.Now}] Connection failed: {errorMsg} (Code: {result})\n");
                throw new IOException("Error connecting to remote share", new Win32Exception(result));
            }
        }

        private string GetUncRoot(string fullPath)
        {
            // Extract root UNC path, e.g., \\192.168.10.100 from \\192.168.10.100\ShareName
            Uri uri = new Uri(fullPath);
            return $"\\\\{uri.Host}";
        }

        public void Dispose()
        {
            WNetCancelConnection2(_networkName, 0, true);
        }

        [DllImport("mpr.dll")]
        private static extern int WNetAddConnection2(NetResource netResource,
            string password, string username, int flags);

        [DllImport("mpr.dll")]
        private static extern int WNetCancelConnection2(string name, int flags, bool force);
    }

    [StructLayout(LayoutKind.Sequential)]
    public class NetResource
    {
        public ResourceScope Scope;
        public ResourceType ResourceType;
        public ResourceDisplaytype DisplayType;
        public int Usage;
        public string LocalName;
        public string RemoteName;
        public string Comment;
        public string Provider;
    }

    public enum ResourceScope : int
    {
        Connected = 1,
        GlobalNetwork,
        Remembered,
        Recent,
        Context
    };

    public enum ResourceType : int
    {
        Any = 0,
        Disk = 1,
        Print = 2,
    }

    public enum ResourceDisplaytype : int
    {
        Generic = 0x0,
        Domain = 0x01,
        Server = 0x02,
        Share = 0x03,
        File = 0x04,
        Group = 0x05,
        Network = 0x06,
        Root = 0x07,
        Shareadmin = 0x08,
        Directory = 0x09,
        Tree = 0x0a,
        Ndscontainer = 0x0b
    }
}
