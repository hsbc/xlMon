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
        public DateTime ProcessStartTime { get; private set; }
        public string LocalIP { get; private set; }

        public DiagnosticInfo()
        {
            p = System.Diagnostics.Process.GetCurrentProcess();
            ProductVersion = p.MainModule.FileVersionInfo.ProductVersion;
            ProcessID = p.Id.ToString();
            MachineName = System.Environment.MachineName;
            Username = System.Environment.UserName;
            UserDomain = System.Environment.UserDomainName;
            ProcessStartTime = p.StartTime;

            using (Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, 0))
            {
                socket.Connect("8.8.8.8", 65530);
                IPEndPoint endPoint = socket.LocalEndPoint as IPEndPoint;
                LocalIP = endPoint.Address.ToString();
            }
        }

        //Properties that will change over time.
        public ulong MilliSecondsProcessRunning
        {
            get
            {
                TimeSpan runningTimeSpan = DateTime.Now - this.ProcessStartTime;
                return Convert.ToUInt64(runningTimeSpan.Ticks / 10000);
            }
        }
    }


}
