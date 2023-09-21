using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace XLMonCOMAddin
{
    static class UDPMessageHelper
    {
        private static JObject CreateCommonElements(DiagnosticInfo d, string sessionId, string evtType)
        {

            JObject msg = new JObject(new JProperty("Evt", evtType))
            {
                { "Time", DateTime.UtcNow.ToString("s") },
                { "SID", sessionId },
                { "UsrNm", Path.Combine(d.UserDomain, d.Username) },
                { "MchNm", d.MachineName },
                { "IP", d.LocalIP },
                { "MchUpTime", DiagnosticInfo.MilliSecondsSinceReboot / 1000 }, 
                { "XLUpTime", d.MilliSecondsProcessRunning / 1000}
            };
            return msg;
        }

        static public string BuildIntroMsg(string sessionId, DiagnosticInfo d)
        {
            //Get the Module version, that is the version of the XLMon dll.
            FileVersionInfo AssemblyVersionInfo = FileVersionInfo.GetVersionInfo(FileHelpers.GetAssemblyFullPath());
            string AssemblyFileVersion = AssemblyVersionInfo.FileVersion;

            JObject introMsg = CreateCommonElements(d, sessionId, "Intro");
            introMsg.Add("XLVer", d.ProductVersion);
            introMsg.Add("XLMonVer", AssemblyFileVersion);

            return introMsg.ToString(Formatting.None);
        }

        static public string BuildStillOpenMsg(List<Tuple<string, ulong>> wkbkTimings, string sessionId, DiagnosticInfo d)
        {
            JObject StillOpenMsg = CreateCommonElements(d, sessionId, "StillOpen");
            if (wkbkTimings.Count > 0)
            {
                JArray jWkBkNode = new JArray();
                foreach (Tuple<string, ulong> wkBkIter in wkbkTimings)
                {
                    JObject jo = new JObject
                    {
                        new JProperty("FP", wkBkIter.Item1),
                        new JProperty("TO", wkBkIter.Item2)
                    };
                    jWkBkNode.Add(jo);
                }
                StillOpenMsg.Add(new JProperty("WkBkData", jWkBkNode));
            }
            return StillOpenMsg.ToString(Formatting.None);
        }

        static public string BuildErrorMsg(string desc, string sessionId, DiagnosticInfo d)
        {
            JObject ErrMsg = CreateCommonElements(d, sessionId, "Error");
            ErrMsg.Add("Desc", desc);
            return ErrMsg.ToString(Formatting.None);
        }

        static public string BuildExitMsg(string sessionId, DiagnosticInfo d)
        {
            return CreateCommonElements(d, sessionId, "Exit").ToString(Formatting.None);
        }

        static public string BuildRegAddinDetailsMsg(string sessionId, DiagnosticInfo d, List<Dictionary<string, string>> registryDetails)
        {
            JObject RegAddinDetailsMsg = CreateCommonElements(d, sessionId, "RegAddinDetails");
            if (registryDetails.Count > 0)
            {
                JArray jRegNode = new JArray();
                foreach (Dictionary<string, string> addinDictionary in registryDetails)
                {
                    JObject jo = new JObject
                    {
                        new JProperty("Type", addinDictionary["AddinType"]),
                        new JProperty("Name", addinDictionary["AddinName"]),
                        new JProperty("Path", addinDictionary["Path"]),
                        new JProperty("LoadBehavior", addinDictionary["LoadBehavior"])
                    };
                    jRegNode.Add(jo);
                }
                RegAddinDetailsMsg.Add(new JProperty("RegDetails", jRegNode));
            }
            return RegAddinDetailsMsg.ToString(Formatting.None);
        }

        static public string BuildXLLMsg(in string sessionId, in DiagnosticInfo d, in Dictionary<String, ModuleDetails> xllDetails)
        {
            JObject XLLAddinDetailsMsg = CreateCommonElements(d, sessionId, "LoadAddins");
            if (xllDetails.Count > 0)
            {
                JArray jRegNode = new JArray();
                foreach (var addinDictionary in xllDetails)
                {
                    JObject jo = new JObject
                    {
                        new JProperty("Type", "XLL"),
                        new JProperty("FN", addinDictionary.Key),
                        new JProperty("DM", addinDictionary.Value.DateModified.ToString("s")),
                        new JProperty("Size", addinDictionary.Value.fileSize),
                        new JProperty("V", addinDictionary.Value.fileVersion)
                    };
                    jRegNode.Add(jo);
                }

                //At this point we could also add COM Dlls and Automation Dlls if we could determine them, from every other random dll in the process space


                XLLAddinDetailsMsg.Add(new JProperty("Details", jRegNode));
            }

            return XLLAddinDetailsMsg.ToString(Formatting.None);
        }
    }
}
