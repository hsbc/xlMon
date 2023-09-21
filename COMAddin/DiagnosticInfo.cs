using System;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;

namespace XLMonCOMAddin
{
    //This class wraps useful information about the process and machine that we might want to use elsewhere.
    public class DiagnosticInfo
    {
        private readonly Process p;   //We use this for some of the properties, so cache it here.

        //Properties that stay the same during the process, set in the contructor, read only.
        public string ProductVersion { get; private set; }
        public string ProcessID { get; private set; }
        public string MachineName { get; private set; }
        public string Username { get; private set; }
        public string UserDomain { get; private set; }
        public bool RunningOnServer { get; private set; }
        public DateTime ProcessStartTime { get; private set; }
        public DateTime LastRebootTime { get; private set; }
        public string LocalIP { get; private set; }

        public DiagnosticInfo()
        {
            p = System.Diagnostics.Process.GetCurrentProcess();
            ProductVersion = p.MainModule.FileVersionInfo.ProductVersion;
            ProcessID = p.Id.ToString();
            MachineName = System.Environment.MachineName;
            Username = System.Environment.UserName;
            UserDomain = System.Environment.UserDomainName;
            RunningOnServer = NativeMethods.IsOS(NativeMethods.OS_ANYSERVER);
            ProcessStartTime = p.StartTime;

            //Calc Reboot Time.
            TimeSpan tsUpTime = TimeSpan.FromSeconds(NativeMethods.GetTickCount64() / 1000);
            LastRebootTime = DateTime.Now.Subtract(tsUpTime);

            using (Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, 0))
            {
                socket.Connect("8.8.8.8", 65530);
                IPEndPoint endPoint = socket.LocalEndPoint as IPEndPoint;
                LocalIP = endPoint.Address.ToString();
            }
        }

        //Properties that will change over time.
        public static ulong MilliSecondsSinceReboot { get { return NativeMethods.GetTickCount64(); } }
        public int NumThreads { get { return p.Threads.Count; } }
        public long WorkingSetMemory { get { return p.WorkingSet64; } }
        public long PrivateMemory { get { return p.PrivateMemorySize64; } }
        public int NumHandles { get { return p.HandleCount; } }
        public int NumGDIResources { get { return NativeMethods.GetGuiResources(p.Handle, 0); } }
        public ulong MilliSecondsProcessRunning
        {
            get
            {
                TimeSpan runningTimeSpan = DateTime.Now - this.ProcessStartTime;
                return Convert.ToUInt64(runningTimeSpan.Ticks / 10000);
            }
        }
    }


    //Microsoft recommends (via Code Analysis) that native calls are held in their own internal class so it's clear to see where these are being made. https://msdn.microsoft.com/library/ms182161.aspx
    internal static class NativeMethods
    {
        [DllImport("User32")]
        public static extern int GetGuiResources(IntPtr hProcess, int uiFlags);

        [DllImport("kernel32.dll")]
        public static extern UInt64 GetTickCount64();

        [DllImport("shlwapi.dll", SetLastError = true, EntryPoint = "#437")]
        public static extern bool IsOS(int os);
        public const int OS_ANYSERVER = 29;
    }

}
