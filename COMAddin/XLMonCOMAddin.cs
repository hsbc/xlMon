using log4net.Util.TypeConverters;
using log4net.Util;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.Net.Sockets;

namespace XLMonCOMAddin
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    [System.Runtime.InteropServices.ProgId("HSBC_XLMonCOMAddin.MyConnect")]
    [System.Runtime.InteropServices.Guid("ffffffff-974d-44a3-8a5e-100000000001")]

    [System.Runtime.InteropServices.ComVisible(true)]
    public class XLMonCOMAddin : Extensibility.IDTExtensibility2
    {
        Microsoft.Office.Interop.Excel.Application AppObj;
        Microsoft.Office.Core.COMAddIn AddinInst;

        private static readonly log4net.ILog filelog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private static readonly log4net.ILog udpLog = log4net.LogManager.GetLogger("udpToServer");

        private static readonly string sessionId = Guid.NewGuid().ToString();   //Constant for this session so that we can join all the info about this Excel.exe session
        private DateTime lastTimeWorkBooksLoggedUTC;                            //Used so that we only iterate the workbooks and send the UDP message at max frequency "WorkbookLogIntervalSeconds"
        private bool SeenFirstError = false;                                    //We use this to ensure that we only send a single UDP Error message for Excel Events, and also after this we do nothing
        private int WorkbookLogIntervalSeconds = 120;                           //Default is 120 seconds, but this can be overriden from the registry
        private readonly DiagnosticInfo diagInfo = new DiagnosticInfo();        //Will be passing to all UDP Message Functions
        private readonly UNCCache UNCLookupCache = new UNCCache();
        private readonly Workbook_Timer workbookTimings = new Workbook_Timer();
        private readonly ModuleScanner xllModScanner = new ModuleScanner();

        public void OnConnection(
            object Application,
            Extensibility.ext_ConnectMode ConnectMode,
            object AddInInst,
            ref Array custom)
        {
            try
            {
                //See if we can get a reg entry that says we show a Msgbox on Start-up, this is designed purely for testing purposes, in case we aren't getting logs we expect.
                bool showIntroBox = ShowIntroBox();
                if (showIntroBox)
                {
                    System.Windows.Forms.MessageBox.Show("Initialising COM Addin" + Environment.NewLine + FileHelpers.GetAssemblyFullPath(), "HSBC XLMon");
                }

                this.AppObj = (Microsoft.Office.Interop.Excel.Application)Application;
                if (this.AddinInst == null)
                {
                    this.AddinInst = (Microsoft.Office.Core.COMAddIn)AddInInst;
                    this.AddinInst.Object = this;
                }

                SetupLog4Net();
                //Read in Frequency
                string sFrequency = GetConfigRegEntry("Frequency", WorkbookLogIntervalSeconds.ToString());
                int.TryParse(sFrequency, out WorkbookLogIntervalSeconds);  //Ignore return value, if it works then we use it, else it has same value as before;
                filelog.Info("Logging every '" + WorkbookLogIntervalSeconds.ToString() + "'");
                filelog.Info("Using Assembly: " + FileHelpers.GetAssemblyFullPath());

                lastTimeWorkBooksLoggedUTC = new DateTime(1970, 1, 1);

                //We are not relying on workbook close to calculate workbook open time as we would get no results if Excel.exe crashed or was terminated

                filelog.Info("About to hook into Excel Events");
                AppObj.SheetActivate += AppObj_SheetActivate;
                AppObj.AfterCalculate += AppObj_AfterCalculate;
                AppObj.WorkbookActivate += AppObj_WorkbookActivate;
                AppObj.SheetSelectionChange += AppObj_SheetSelectionChange;
                AppObj.WorkbookOpen += AppObj_WorkbookOpen;

                string logMsg = UDPMessageHelper.BuildIntroMsg(sessionId, diagInfo);
                SendUDP(logMsg);

                if (showIntroBox)
                {
                    System.Windows.Forms.MessageBox.Show("Finished Initialising COM Addin", "HSBC XLMon");
                }

                // Extract registry details
                string strRegistryDetails = RegistryHelper.ReportRegistryDetails(sessionId, diagInfo);
                SendUDP(strRegistryDetails);
            }
            catch (Exception ex)
            {
                SeenFirstError = true;
                // This is one of the few cases we log what we send over UDP differently to what we write to the file log.
                //Don't want to be sending exception string over UDP.
                filelog.Error(ex.ToString());
                udpLog.Error(UDPMessageHelper.BuildErrorMsg("Startup", sessionId, diagInfo));
            }
        }

        public void OnDisconnection(
            Extensibility.ext_DisconnectMode RemoveMode,
            ref Array custom)
        {
            if (!SeenFirstError)
            {
                try
                {
                    //Don't care how long since we last sent message, we want to log workbooks, and XLLs
                    filelog.Info("Starting Disconnection. Doing final write of XLS Times and Loaded XLL Add-Ins ");
                    LogWorkBooks();
                    LogXLLAddins(true);

                    string logMsg = UDPMessageHelper.BuildExitMsg(sessionId, diagInfo);
                    SendUDP(logMsg);

                    filelog.Info("About to unhook into Excel Events");
                    AppObj.WorkbookOpen -= AppObj_WorkbookOpen;
                    AppObj.SheetSelectionChange -= AppObj_SheetSelectionChange;
                    AppObj.WorkbookActivate -= AppObj_WorkbookActivate;
                    AppObj.AfterCalculate -= AppObj_AfterCalculate;
                    AppObj.SheetActivate -= AppObj_SheetActivate;

                    filelog.Info("Finished OnDisconnection");
                }
                catch (Exception ex)
                {
                    filelog.Error(ex.ToString());
                    udpLog.Error(UDPMessageHelper.BuildErrorMsg("Shutdown", sessionId, diagInfo));
                }

                this.AppObj = null;
                this.AddinInst = null;
            }
        }

        //Required Stubs so we are fully implementing the interface
        public void OnAddInsUpdate(ref Array custom) {}
        public void OnStartupComplete(ref Array custom) {}
        public void OnBeginShutdown(ref Array custom) {}

        private static bool ShowIntroBox()
        {
            try
            {
                //If the reg key exists and has any value, i.e. length > 0 then we return true.
                object oRegResult = Registry.GetValue(@"HKEY_CURRENT_USER\Software\HSBC\XLMon", "ShowIntroDialog", "");
                return oRegResult != null && oRegResult.ToString().Length > 0;
            }
            catch
            {
                //If anything goes wrong here in theory we would want to flag it with a message box, but we can't worry the user.
                return false;
            }
        }

        private void AppObj_SheetActivate(object Sh)
        {
            HandleEvent("SheetActivate");
        }

        private void AppObj_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            //We must put a try catch around any event from Excel. If something goes wrong then SeenFirstError is set, we try and send a UDP message, then go super quiet
            if (!SeenFirstError)
            {
                try
                {
                    //We have to handle workbooks that have been opened from a template differently, these haven't been saved and have .FullName that is not a path
                    string wkPath = Wb.Path; //This will be zero length if the workbook has never been saved, i.e a new workbook.
                    if (wkPath.Length > 0)
                    {
                        string WbFullName = Wb.FullName;
                        workbookTimings.WorkbookOpened(UNCLookupCache.GetUNC(WbFullName));
                    }
                    HandleEvent("WorkbookOpen");
                }
                catch (Exception ex)
                {
                    if (!SeenFirstError)  //Make sure we only send this once, we don't want to hear about 100's of error messages, the first is good enough.
                    {
                        SeenFirstError = true;
                        filelog.Error(ex.ToString());
                        udpLog.Error(UDPMessageHelper.BuildErrorMsg("_ApplicationObject_WorkbookOpen", sessionId, diagInfo));
                    }
                }
            }
        }

        private void AppObj_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
        {
            HandleEvent("SheetSelectionChange");
        }

        private void AppObj_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook Wb)
        {
            HandleEvent("WorkbookActivate");

            //This looks wrong, but the COM Addins implemented in C# seem to require us to do a memory clean up of any COM References we have manaually and it has to be called twice.
            //If we don't do this the workbooks stay visible in the VBA Editor and we are likely to get close down problems. This event seems like the best one to put it in.
            //We don't want this called too often.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void AppObj_AfterCalculate()
        {
            HandleEvent("AfterCalculate");
        }

        private bool TimeToLogWorkbooks()
        {
            DateTime now = DateTime.UtcNow;
            TimeSpan ts = now - lastTimeWorkBooksLoggedUTC;
            return Convert.ToInt32(ts.TotalSeconds) > WorkbookLogIntervalSeconds;
        }

        private void LogWorkBooks()
        {
            List<string> currentlyOpenWorkbooks = new List<string>();
            foreach (Excel.Workbook wk in this.AppObj.Workbooks)
            {
                string wkPath = wk.Path; //This will be zero length if the workbook has never been saved, i.e a new workbook.
                                         //We are only tracking opened work books, not new temp ones.
                if (wkPath.Length > 0)
                {
                    currentlyOpenWorkbooks.Add(UNCLookupCache.GetUNC(wk.FullName));
                }
            }

            //Send all this information to the Timings Class so it can work out what to send.
            workbookTimings.CurrentWorkbooksOpen(currentlyOpenWorkbooks);

            List<Tuple<string, ulong>> wkBkTimings = workbookTimings.GetWorkBookTimes();
            string logMsg = UDPMessageHelper.BuildStillOpenMsg(wkBkTimings, sessionId, diagInfo);

            SendUDP(logMsg);
        }

        private static void SendUDP( string logMsg )
        {
            filelog.Info("Sending : " + logMsg);
            udpLog.Info(logMsg);
        }

        private void HandleEvent(string EventName)
        {
            if (!SeenFirstError)
            {
                try
                {
                    if (TimeToLogWorkbooks())
                    {
                        LogWorkBooks();
                        LogXLLAddins(false);
                        lastTimeWorkBooksLoggedUTC = DateTime.UtcNow;
                    }
                }
                catch (Exception ex)
                {
                    if (!SeenFirstError)  //Make sure we only send this once, we don't want to hear about 100's of error messages, the first is good enough.
                    {
                        SeenFirstError = true;
                        filelog.Error(ex.ToString());
                        udpLog.Error(UDPMessageHelper.BuildErrorMsg(EventName, sessionId, diagInfo));
                    }
                }
            }
        }


        private void SetupLog4Net()
        {
            GetServerAndPort(out string UDPServerAddress, out int UDPServerPort);

            log4net.Util.TypeConverters.ConverterRegistry.AddConverter(typeof(IPAddress), new IPAddressPatternConverter());
            log4net.Util.TypeConverters.ConverterRegistry.AddConverter(typeof(int), new NumericConverter());

            log4net.GlobalContext.Properties["XLM_ProcessID"] = Process.GetCurrentProcess().Id.ToString();
            log4net.GlobalContext.Properties["XLM_sessionId"] = sessionId;
            log4net.GlobalContext.Properties["XLM_tempFolder"] = Path.GetTempPath();
            log4net.GlobalContext.Properties["XLM_ServerName"] = UDPServerAddress;
            log4net.GlobalContext.Properties["XLM_ServerPort"] = UDPServerPort;

            FileInfo fi = new FileInfo(Path.Combine(FileHelpers.GetAssemblyFolder(), "log4net.config"));
            log4net.Config.XmlConfigurator.Configure(fi);

            filelog.Info("File Logging Initialised");
            filelog.Info("Sending Usage Data to Server '" + UDPServerAddress + "' Port '" + UDPServerPort + "'");
        }

        private void GetServerAndPort(out string Server, out int PortNum)
        {
            //Insert your defaults here if you want. Values are overriden from registry.
            const int defaultPort = 40042;
            const string defaultServer = "defaultServername.domain";

            Server = GetConfigRegEntry("Server", defaultServer);
            string sPortNum = GetConfigRegEntry("Port", defaultPort.ToString());

            if (!int.TryParse(sPortNum, out PortNum))
            {   //Don't log this yet, we'll log the results later
                PortNum = defaultPort;
            }
        }

        private static string GetConfigRegEntry(string KeyName, string DefaultValue)
        {
            //We use this to get our settings under XLMon. The standard default mechanism doesn't work if the higher level XLMon reg key doesn't exist,
            //so we still need to check for null return value.

            object oRegResult = Registry.GetValue(@"HKEY_CURRENT_USER\Software\HSBC\XLMon", KeyName, DefaultValue);
            if (oRegResult == null)
            {
                return DefaultValue;
            }
            else
            {
                return oRegResult.ToString();
            }
        }


        private void LogXLLAddins(bool waitForResult)
        {
            if (xllModScanner.GetXLLModules(out Dictionary<string, ModuleDetails> xllDetails, waitForResult))
            {
                string strXLLDetailsMsg = UDPMessageHelper.BuildXLLMsg(sessionId, diagInfo, xllDetails);
                SendUDP(strXLLDetailsMsg);
            }
        }
                
        public class NumericConverter : IConvertFrom
        {
            public NumericConverter() { }

            public Boolean CanConvertFrom(Type sourceType)
            {
                return typeof(String) == sourceType;
            }

            public Object ConvertFrom(Object source)
            {
                String pattern = (String)source;
                PatternString patternString = new PatternString(pattern);
                String value = patternString.Format();
                return Int32.Parse(value);
            }
        }

        public class IPAddressPatternConverter : IConvertFrom
        {
            public IPAddressPatternConverter() { }

            public Boolean CanConvertFrom(Type sourceType)
            {
                return typeof(String) == sourceType;
            }

            public Object ConvertFrom(Object source)
            {
                String pattern = (String)source;
                PatternString patternString = new PatternString(pattern);
                String value = patternString.Format();

                foreach (IPAddress IPA in Dns.GetHostAddresses(value))
                {
                    if (IPA.AddressFamily == AddressFamily.InterNetwork)
                    {
                        return IPA;
                    }
                }
                return null;
            }
        }
    }
}