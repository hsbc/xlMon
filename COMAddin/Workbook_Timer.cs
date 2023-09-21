using System;
using System.Collections.Generic;

namespace XLMonCOMAddin
{
    //This class is used to calculate the time a workbook is open. There is no AfterClose Event in Excel (we could simulate one based on Sheet Activate).
    //We also don't want to trust WorkBook Close because if Excel crashes or is terminated then we get no data

    public class Workbook_Timer
    {
        //Stores the info about when the workbook opened. For all workbooks that have been opened in the session
        private readonly Dictionary<string, WorkbookMeta> wbMetaMap = new Dictionary<string, WorkbookMeta>();

        public void WorkbookOpened(string wbName)
        {
            if (!wbMetaMap.ContainsKey(wbName))
            {
                WorkbookMeta wbMeta = new WorkbookMeta();
                wbMetaMap.Add(wbName, wbMeta);
            }

            // the meta class will handle the scenario where the workbook is closed, already open, etc
            wbMetaMap[wbName].Open();
        }

        public void CurrentWorkbooksOpen(List<string> currentlyOpenWBs)
        {
            if (currentlyOpenWBs != null)
            {
                // check all workbooks currently open, by comparing against previous list we can determine which ones have been closed.
                foreach (string wb in wbMetaMap.Keys)
                {
                    if (wbMetaMap[wb].IsOpen() && !currentlyOpenWBs.Contains(wb))
                    {
                        wbMetaMap[wb].Close();
                    }
                }

                // execute open for all the active wbs - no impact on timing, if already opened
                foreach (string openWB in currentlyOpenWBs)
                {
                    WorkbookOpened(openWB);
                }
            }
        }

        public List<Tuple<string, ulong>> GetWorkBookTimes()
        {
            //send out all the updates from the publishing queue.
            List<Tuple<string, ulong>> results = new List<Tuple<string, ulong>>();
            foreach (string wb in wbMetaMap.Keys)
            {
                results.Add(new Tuple<string, ulong>(wb, wbMetaMap[wb].GetCurrentUptime()));
            }
            return results;
        }

    }
}
