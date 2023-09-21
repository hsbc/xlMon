using Microsoft.Win32;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace XLMonCOMAddin
{
    static class RegistryHelper
    {
        static public string ReportRegistryDetails(in string sessionId, in DiagnosticInfo d)
        {
            List<Dictionary<string, string>> registryDetails = new List<Dictionary<string, string>>();

            // 1. Standard OPEN Excel Reg Keys
            string DefaultValue = "DefaultValue";
            for (int i = 0; i <= 50; i++)
            {
                Dictionary<string, string> addinDetails = new Dictionary<string, string>();
                string sName = (i == 0) ? "OPEN" : "OPEN" + i;
                object oRegResult = Registry.GetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Options", sName, DefaultValue);
                if (oRegResult == null || oRegResult.ToString() == DefaultValue)
                    break;
                addinDetails.Add("AddinType", "HKCU_ExcelOptions");
                addinDetails.Add("AddinName", sName);

                SplitExcelOptions(oRegResult.ToString(), out string sPath, out string sLoadBehavior);

                addinDetails.Add("Path", sPath);
                addinDetails.Add("LoadBehavior", sLoadBehavior);
                registryDetails.Add(addinDetails);
            }

            // 2. COM Add-ins Computer\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\Excel\AddIns
            GetCOMAddinDetails(RegistryHive.CurrentUser, ref registryDetails);

            // 3. COM Add-ins Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\Excel\AddIns
            GetCOMAddinDetails(RegistryHive.LocalMachine, ref registryDetails);

            return UDPMessageHelper.BuildRegAddinDetailsMsg(sessionId, d, registryDetails);
        }

        static private void SplitExcelOptions(in string inputStr, out string strPath, out string strPrefix)
        {
            /* Example for below regex split
             * inputStr - "/R \"C:\\Program Files (x86)\\Microsoft Office\\Office16\\Library\\Analysis\\ANALYS32.XLL\""
             * strPrefix - /R
             * strPath - C:\\Program Files (x86)\\Microsoft Office\\Office16\\Library\\Analysis\\ANALYS32.XLL\
             */
            var re = new Regex("(?<=\")[^\"]*(?=\")|[^\" ]+");
            var matches = re.Matches(inputStr);
            strPath = "";
            strPrefix = "";
            int count = 0;
            foreach (Match match in matches)
            {
                count++;
                if (matches.Count == count)
                {
                    strPath = match.Value;
                }
                else
                {
                    strPrefix += match.Value;
                }
            }
            return;
        }

        static private void GetCOMAddinDetails(in RegistryHive nBaseKey, ref List<Dictionary<string, string>> registryDetails)
        {
            RegistryKey baseKey = RegistryKey.OpenBaseKey(nBaseKey, RegistryView.Default);
            if (baseKey != null)
            {
                RegistryKey addInsKey = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Office\Excel\AddIns");
                // If Key is found for AddIns. Extract all list of addins and get the LoadBehavior
                if (addInsKey != null)
                {
                    string[] addIns = addInsKey.GetSubKeyNames();
                    foreach (string addIn in addIns)
                    {
                        Dictionary<string, string> addinDetails = new Dictionary<string, string>();
                        RegistryKey addinKey = addInsKey.OpenSubKey(addIn);
                        if (addinKey != null)
                        {
                            object oRegResult = addinKey.GetValue("LoadBehavior");
                            addinDetails.Add("AddinType", (nBaseKey == RegistryHive.CurrentUser) ? "HKCU_CLSID" : "HKLM_CLSID");
                            addinDetails.Add("AddinName", addIn);
                            addinDetails.Add("Path", "");
                            addinDetails.Add("LoadBehavior", (oRegResult != null) ? oRegResult.ToString() : "");
                            registryDetails.Add(addinDetails);
                            addinKey.Close();
                        }
                    }
                    addInsKey.Close();
                }
                baseKey.Close();
            }
        }
    }
}
