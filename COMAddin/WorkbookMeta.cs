using System;

namespace XLMonCOMAddin
{
    class WorkbookMeta
    {
        enum WBSTATUS { INITIALIZED, OPEN, CLOSE };
        WBSTATUS currentStatus;
        DateTime openTime;
        ulong totalUpTime;

        public WorkbookMeta()
        {
            totalUpTime = 0;
            currentStatus = WBSTATUS.INITIALIZED;
        }

        public void Open()
        {
            if (currentStatus != WBSTATUS.OPEN)
            {
                currentStatus = WBSTATUS.OPEN;
                openTime = DateTime.Now;
            }
        }

        public void Close()
        {
            if (currentStatus == WBSTATUS.OPEN)
            {
                currentStatus = WBSTATUS.CLOSE;

                TimeSpan timeElapsed = DateTime.Now - openTime;
                totalUpTime += Convert.ToUInt64(timeElapsed.TotalSeconds);
            }
        }

        public bool IsOpen()
        {
            return (currentStatus == WBSTATUS.OPEN);
        }


        public ulong GetCurrentUptime()
        {
            ulong uptime = totalUpTime;
            if (IsOpen())
            {
                TimeSpan timeElapsed = DateTime.Now - openTime;
                uptime += Convert.ToUInt64(timeElapsed.TotalSeconds);
            }
            // if closed then totalUpTime is all you have
            return uptime;
        }
    }

}
