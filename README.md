# xlMon
 
## Summary
xlMon is a C# Excel COM Add-In for Windows that monitors an Excel session and sends metrics over UDP for a backend to analyse. xlMon collects information about the add-ins that are configured to run on start-up, general information about the Excel session and machine, as well as the full UNC path of workbooks opened and the length of time that they have been open.

You could look at the times that workbooks were open, the patterns around when they are opened etc to determine how important they are.

The repo consists of a single Visual Studio solution, containing two C# projects.
1) A COM Add-In that is installed and run inside the Excel process, that needs to be deployed and registered on the user’s machine.
2) A simple UDP Listener that can be used to test to show that the solution is working as expected, that doesn't need to be deployed.

The xlMon COM Add-In can be installed on a machine without admin rights as it can be configured using on HKCU registry entries, which gives the user the ability to disable it, or HKLM entries which means the user cannot disable\remove it.


## Local Logging
xlMon writes it's log files to %TEMP%\xlmon. Each Excel process will have a separate filename, a combination of process ID and a generated GUID.
Some basic initialisation is written, including details of the remote server and port the UDP Messages are sent to.
Any UDP messages that are sent remotely are also included in the local log.


## Remote Logging
xlMon sends UDP Messages in JSON. There are a few different types listed below, and identified by the Evt field in the JSON.

### Intro
This message is only sent once per Excel process, when Excel first starts up.

```json
{
    "Evt": "Intro",
    "Time": "2023-03-21T23:32:12",
    "SID": "1fa6eae9-ab0a-4806-8c49-1631ce8f7160",
    "UsrNm": "USERNAME_IS_HERE",
    "MchNm": "MACHINE_NAME_HERE",
    "IP": "XXX.XXX.XXX.XXX",
    "MchUpTime": 23947,
    "XLUpTime": 1,
    "XLVer": "16.0.5387.1000",
    "XLMonVer": "1.6.0.0"
}
```

**SID** is a unique ID for the Excel process that is generating the messages. All other messages from the same process will have the same SID,
**UsrName** - The Window Login, **MchNm** - The Machine Name, **IP** - The IP Address, **MchUpTime** - Number of seconds since machine last rebooted, **XLUptime** - Number of seconds the Excel process has been alive for, **XLVer** - The Product Version of Excel, **XLMonVer** - The version of XLMon that has sent the data

### StillOpen
This message is sent periodically. Each message from the same process supersedes the last. This means that if you group by **SID** you can discard all the previous StillOpen message sent, as the content is cumulative. We send periodically rather than send once when Excel closes down, because often Excel closes or the user End Tasks it, in those circumstances we wouldn't get the information.

We include most of the fields from the intro message and a list of WorkBook FullPaths (**FP**) and Total Time they have been open (**TO**).

```json
{
    "Evt": "StillOpen",
    "Time": "2023-03-21T23:31:37",
    "SID": "66e1eac7-efb7-48d1-abdf-a17fd4110064",
    "UsrNm": "USERNAME_IS_HERE",
    "MchNm": "MACHINE_NAME_HERE",
    "IP": "XXX.XXX.XXX.XXX",
    "MchUpTime": 52235,
    "XLUpTime": 434,
    "WkBkData":
    [{
        "FP": "c:\\Temp\\abc.xlsx",
        "TO": 4315
    },
    {
        "FP": "C:\\Temp\\SUMMARY.xlsx",
        "TO": 1939
    },
    {
        "FP": "C:Temp\\DEF.xlsm",
        "TO": 1057
    },
    {
        "FP": "c:\\Temp\\BARNEY.xlsx",
        "TO": 642
    },
    {
        "FP": "C:\\Temp\\ABC.xlsx",
        "TO": 145
    }]
}
```

### RegAddinDetails
This message is sent on start-up once. It lists the XLA, COM and automation Add-ins that are configured. It will list the XLAs and Automations Add-Ins that are configured to load, and also all the information that we have on COM Add-Ins and whether they are loaded at start-up or not.
```json
{
    "Evt": "RegAddinDetails",
    "Time": "2023-03-21T22:36:01",
    "SID": "73ea30cd-4b1f-490d-ae4b-e7fdeb27fa38",
    "UsrNm": "USERNAME_IS_HERE",
    "MchNm": "MACHINE_NAME_HERE",
    "IP": "XXX.XXX.XXX.XXX",
    "MchUpTime": 13788,
    "XLUpTime": 4,
    "RegDetails": [{
        "Type": "HKCU_CLSID",
        "Name": "ADXXLForm",
        "Path": "",
        "LoadBehavior": ""
    },
    {
        "Type": "HKCU_CLSID",
        "Name": "ApplicationPrintAddIn.ExcelAddIn",
        "Path": "",
        "LoadBehavior": "0"
    },
    {
        "Type": "HKCU_CLSID",
        "Name": "HSBC_XLMonCOMAddin.MyConnect",
        "Path": "",
        "LoadBehavior": "3"
    },
    {
        "Type": "HKLM_CLSID",
        "Name": "ExcelPlugInShell.PowerMapConnect",
        "Path": "",
        "LoadBehavior": "2"
    },
    {
        "Type": "HKLM_CLSID",
        "Name": "MSIP.ExcelAddin",
        "Path": "",
        "LoadBehavior": "3"
    },
    {
        "Type": "HKLM_CLSID",
        "Name": "TFCOfficeShim.Connect.15",
        "Path": "",
        "LoadBehavior": "3"
    },
    {
        "Type": "HKLM_CLSID",
        "Name": "VS15ExcelAdaptor",
        "Path": "",
        "LoadBehavior": "3"
    }]
}
```

### LoadAddins
This message gets sent at half the frequency of the StillOpen messages and like them we only need to look at the latest message, they are cumulative. We can discard all previous messages from the same SID of this type. The reason we monitor for these is that lots of code within HSBC loads XLLs dynamically. There is no event to respond to, therefore we have to periodically scan the process space looking for any newly loaded modules.

There is an argument that we should make this more generic and also monitor for COM dlls and Automation dlls that have been loaded since we started up, but this would mean trying to determine which dlls in the memory space where those types of files, not just normally dlls. If an XLL is already configured to load at start-up we will currently get a duplication, but not much we can do about this as it's hard to tell if an XLL has been loaded at start-up or dynamically.

```json
{
    "Evt":"LoadAddins",
    "Time":"2023-09-12T16:49:01",
    "SID":"5780610c-eec8-4da9-8377-153540adaa80",
    "UsrNm": "USERNAME_IS_HERE",
    "MchNm": "MACHINE_NAME_HERE",
    "IP": "XXX.XXX.XXX.XXX",
    "MchUpTime":6501,
    "XLUpTime":61,
    "Details":
    [{
        "Type":"XLL",
        "FN":"C:\\ProgramData\\Interface.xll",
        "DM":"2023-09-08T17:17:08",
        "Size":751104,
        "V":"0.33.9.1"
    },
    {
        "Type":"XLL",
        "FN":"C:\\ProgramData\\SpiritToolkit.xll",
        "DM":"2023-09-08T17:39:08",
        "Size":6643200,
        "V":null
    }]
}
```


### Exit
Sent when Excel closes down. Every session should have a close down message, but it likely will not get sent if Excel is End Tasked, or crashes. So can't be relied on for any reporting, with the exception or detecting the ratio of Excel crashes to clean close downs.
```json
{
    "Evt": "Exit",
    "Time": "2023-03-22T07:52:37",
    "SID": "eb21f5d0-3c2d-4197-8ad4-8e024f4dea91",
    "UsrNm": "USERNAME_IS_HERE",
    "MchNm": "MACHINE_NAME_HERE",
    "IP": "XXX.XXX.XXX.XXX",
    "MchUpTime": 146,
    "XLUpTime": 20
}
```




