using ExtensionMethods;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace XLMonCOMAddin
{
    public struct ModuleDetails
    {
        public DateTime DateModified;
        public long fileSize;
        public string fileVersion;

        public ModuleDetails(DateTime DateModified, long fileSize, string fileVersion)
        {
            this.DateModified = DateModified;
            this.fileSize = fileSize;
            this.fileVersion = fileVersion;
        }
    }

    public class ModuleScanner
    {
        //Class that we can call to get the XLL Modules without blocking.
        //Each time we'll check to see if the Task has completed, if it has we return True with the data,
        //Otherwise we return False, but before doing so we check the task.
        //If it's still running then fine, leave it
        //If it's not running then (re)start it

        //Task Cancellation -   Not worrying about Task Cancellation, the first time this is called the task might take a little time to read the file details,
        //                      but each subsequent call will only get file data for the new xlls loaded.

        private Task getXLLModulesTask;
        private readonly object _lock = new object();

        //No Need to lock this because we only access it after the Task has completed. Each time we get data we add new results to XLLDetails
        private readonly Dictionary<string, ModuleDetails> XLLDetails = new Dictionary<string, ModuleDetails>();

        public bool GetXLLModules(out Dictionary<string, ModuleDetails> retXLLDetails, bool waitForResult)
        {
            lock (_lock)
            {
                retXLLDetails = null;

                if ((getXLLModulesTask != null) && getXLLModulesTask.IsCompleted)
                {
                    //The Simple case, we have some results.
                    //Get the data and return to the caller - Don't need to lock XLLDetails because the task has finished, and there is an entry lock around this method.

                    //Wipe out Task, we are using this as a signal to say next time we are called we need to recreate it.
                    getXLLModulesTask = null;
                    retXLLDetails = new Dictionary<string, ModuleDetails>(XLLDetails);
                    return true;
                }
                else if ((getXLLModulesTask != null) && getXLLModulesTask.IsFaulted)
                {
                    throw new ApplicationException("Panic");
                }
                else
                {
                    if (!waitForResult) //Task is still running and or is null and needs to be started
                    {
                        if (getXLLModulesTask == null)
                        {
                            getXLLModulesTask = Task.Factory.StartNew(() => GetXllModulesImp(), CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default);
                        }
                        return false;
                    }
                    else
                    {
                        // We are going to wait for the result.
                        if (getXLLModulesTask != null)
                        {
                            getXLLModulesTask.Wait();
                            getXLLModulesTask = null;
                            GetXllModulesImp(); //Call it again, to catch the edge case where we started the task before we loaded the xll, but on Exit we want to make sure we get them all.
                        }
                        else
                        {
                            //There is no running task and we have been asked to wait for it, so just run it
                            GetXllModulesImp();
                        }
                        retXLLDetails = new Dictionary<string, ModuleDetails>(XLLDetails);
                        return true;
                    }

                }

            }

        }

        private void GetXllModulesImp()
        {
            //Get the Loaded XLL Modules.

            List<string> latestXLLs = new List<string>();
            foreach (ProcessModule m in Process.GetCurrentProcess().Modules)
            {
                string fileName = m.FileName ?? string.Empty;
                if (fileName.Right(4).ToLower().EndsWith(".xll"))
                {
                    //Found an XLL that we are interested in.
                    latestXLLs.Add(fileName);
                }
            }

            //Now go through collection, see if we have any new items, if we do get their Module Details, insert into Temp
            foreach (string iFile in latestXLLs)
            {
                if (!XLLDetails.ContainsKey(iFile))
                {
                    //We've found a new dll that we didn't know about.
                    FileInfo fi = new FileInfo(iFile);
                    XLLDetails[iFile] = new ModuleDetails(fi.LastWriteTime, fi.Length, FileVersionInfo.GetVersionInfo(iFile).FileVersion);
                }
            }
        }
    }
}


namespace ExtensionMethods
{
    public static class MyExtensions
    {
        public static string Right(this string sValue, int iMaxLength)
        {
            //Check if the value is valid
            if (string.IsNullOrEmpty(sValue))
            {
                //Set valid empty string as string could be null
                sValue = string.Empty;
            }
            else if (sValue.Length > iMaxLength)
            {
                //Make the string no longer than the max length
                sValue = sValue.Substring(sValue.Length - iMaxLength, iMaxLength);
            }

            //Return the string
            return sValue;
        }
    }
}