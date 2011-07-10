' ================================================================================================== 
' ================================================================================================== 
' "Monitor Exchange Servers", by Harold "Waldo" Grunenwald (harold.grunenwald@gmail.com) 
' 
' Script Description:  Monitor the status of the Exchange Servers.  Checking: 
'        Pingable 
'        Services Running 
'        Connector Status 
'        DataStores Mounted 
'        Transaction Log Drives more than 50% free 
'  
' Level of Fun writing this:  High 
' ================================================================================================== 
' Intention:  Monitor the Exchange servers. 
' Usage:  Run as a Scheduled Task, cycling every 5 minutes or so 
'            Sends email alerts on problems 
' ================================================================================================== 
' Credits: 
' Some Code From: 
' http://www.microsoft.com/technet/scriptcenter/scripts/hardware/monitor/hwmovb07.mspx?mfr=true 
' http://blogs.technet.com/b/heyscriptingguy/archive/2004/10/13/how-can-i-determine-the-percentage-of-free-space-on-a-drive.aspx 
' ADExplorer 
' Scriptomatic2 
'  
' HUGE THANKS to the Microsoft Scripting Team (especially Ed Wilson for putting 
' with my pedantic emails) as well as Bryce Cogswell and Mark Russinovich from the  
' SysInternals team 
' ================================================================================================== 
' ================================================================================================== 
 
'VBScript uses 0-indexed arrays, so the arrServers is (n-1) 
dim arrServers(5)            ' Array to hold the computers to check 
arrServers(0) = "mailserver1"        ' First Computer 
arrServers(1) = "mailserver2"        ' Second Computer 
arrServers(2) = "mailserver3"        ' Do I really need to keep enumerating? 
arrServers(3) = "mailserver4"        ' Ok, maybe I do.... 
arrServers(4) = "mailserver5"        ' Now you're just being silly 
arrServers(5) = "mailserver6"        ' Yup.  You got me... 
'arrServers(6) = "blah"' to get a no-ping result 
 
 
strServerProperties                        = "Name,msExchESEParamLogFilePath" 
Const wbemFlagReturnImmediately            = &h10 
Const wbemFlagForwardOnly                = &h20 
Const ADS_SCOPE_SUBTREE                    = 2 
Set objConnection                        = CreateObject("ADODB.Connection") 
Set objCommand                            = CreateObject("ADODB.Command") 
objConnection.Provider                    = "ADsDSOObject" 
objConnection.Open "Active Directory Provider" 
Set objCommand.ActiveConnection            = objConnection 
objCommand.Properties("Page Size")        = 1000 
objCommand.Properties("Searchscope")    = ADS_SCOPE_SUBTREE  
 
 
set ExchangeInterface                    = CreateObject("CDOEXM.ExchangeServer") 
set StorageGroupInterface                = CreateObject("CDOEXM.StorageGroup") 
set MailStoreInterface                    = CreateObject("CDOEXM.MailboxStoreDB") 
 
 
alertMessage = "" 
 
 
' ---Check on dem boxen!--- 
For each strComputer in arrServers 
    If (IsAlive(strComputer)) Then 
        '---If it pings, check for...--- 
        Set objWMIService = GetObject("winmgmts:" _ 
            & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
             
            '---Connect to Exchange's WMI Provider--- 
            Set obj_EX_WMIService = GetObject("winmgmts:" _ 
                & "{impersonationLevel=impersonate}!\\" &  _ 
                    strComputer & "\root\cimv2\Applications\Exchange") 
             
            '---Check Services--- 
            Set col_EX_Server_State_Items = obj_EX_WMIService.ExecQuery _ 
                ("Select * from ExchangeServerState") 
 
            For Each obj_EX_Server_State_Item in col_EX_Server_State_Items 
                ' This reports on everything in the cluster, so we're having 
                ' each machine report it's own status, hence the 'if [name = self]' 
                if (lcase(obj_EX_Server_State_Item.Name) = lcase(strComputer)) then 
                    '---Connect to WMI?--- 
                    if (obj_EX_Server_State_Item.Unreachable) then 
                        alertMessage = alertMessage  & "Cannot connect to WMI on " _ 
                            & strComputer & VbCrLf 
                    end if 
                     
                    '---Cluster State--- 
                    if (obj_EX_Server_State_Item.ClusterState <> 1) then 
                        alertMessage = alertMessage & strComputer _ 
                            & " - Cluster Issues" & VbCrLf 
                    end if 
                     
                    '---In Maintenance Mode?--- 
                    if (obj_EX_Server_State_Item.ServerMaintenance) then 
                        alertMessage = alertMessage & strComputer _ 
                            & " in Maintenance Mode" & VbCrLf 
                    end if 
                     
                    '---Server State--- 
                    if (obj_EX_Server_State_Item.ServerState <> 1) then 
                        alertMessage = alertMessage & strComputer _ 
                            & " Server state: " _ 
                            & obj_EX_Server_State_Item.ServerStateString & VbCrLf 
                    end if 
                     
                    '---Services State--- 
                    if (obj_EX_Server_State_Item.ServicesState <> 1) then 
                        alertMessage = alertMessage & strComputer _ 
                            & " Services state: " _ 
                            & obj_EX_Server_State_Item.ServicesStateString & VbCrLf 
                    end if 
                end if 
            Next 
             
             
            '---Check Connectors--- 
            Set col_EX_Connector_Items = obj_EX_WMIService.ExecQuery _ 
                ("Select Name,IsUp from ExchangeConnectorState") 
            For Each obj_EX_Connector_Item in col_EX_Connector_Items 
                if (obj_EX_Connector_Item.IsUp = "False") then 
                    alertMessage = alertMessage & obj_EX_Connector_Item.Name _ 
                        & " is Down" & VbCrLf 
                end if 
            Next 
             
             
            '---Check Store Mounts--- 
            ExchangeInterface.DataSource.Open strComputer 
            ' examine the SGs; for each SG, list it's associated stores and paths 
            for each sg in ExchangeInterface.StorageGroups 
                'Exclude the Recovery Storage Groups since CDOEXM doesn't work on them 
                if (1 > (instr(sg,"Recovery"))) then 
                    'parse out the shortname from the DN without querying LDAP 
                    strStorageGroup = mid(sg,4,((instr(sg,",")) - 4)) 
                    StorageGroupInterface.DataSource.Open sg 
                    ' The next line will give you the TLOG file paths for all EX servers, 
                    '    but I can't find how to break it down to just the TLOGs 
                    '    for each particular server 
                    'WScript.Echo StorageGroupInterface.LogFilePath 
                    count = 0 
                    for each mailDB in StorageGroupInterface.MailboxStoreDBs 
                        count = count + 1 
                        MailStoreInterface.DataSource.Open mailDB 
                        Select Case MailStoreInterface.Status 
                            Case "0" 
                                dbStatus = "Running" 
                            Case "1" 
                                dbStatus = "Not Running" 
                            Case "2" 
                                dbStatus = "Mounting - Reserved" 
                            Case "3" 
                                dbStatus = "Dismounting - Reserved" 
                        End Select 
                         
                        if (MailStoreInterface.Status <> 0) then 
                            if (1 > (instr(MailStoreInterface.name,"Temp"))) then 
                                alertMessage = alertMessage & strStorageGroup _ 
                                    & VbCrLf & VbTab & MailStoreInterface.name _ 
                                        & ": " & dbStatus & VbCrLf  
                            end if 
                        end if 
 
                    next 
                end if 
            next 
             
             
            '---Check TLOG Drive Util--- 
            strLDAPOU = "CN=" & ucase(strComputer) _ 
                & ",CN=Servers,CN=MyAdminGroup,CN=Administrative Groups,CN=MyCompanyName,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=subdomain,DC=domain,DC=com" 
            objCommand.CommandText = _ 
                "Select " & strServerProperties &" from 'LDAP://" & strLDAPOU & "' " & _ 
                "Where objectClass = 'msExchStorageGroup' and name <> 'Recovery Storage Group'" 
                ' We Exclude the Recovery Storage Group 
            Set objRecordSet = objCommand.Execute 
            objRecordSet.MoveFirst 
            Do Until objRecordSet.EOF 
                strStorageGroup    = objRecordSet.Fields("Name").Value 
                strLogFilePath    = objRecordSet.Fields("msExchESEParamLogFilePath").Value 
                strLogFileDrive = left(strLogFilePath,2) 
                Set colDisks = objWMIService.ExecQuery _ 
                    ("Select * from Win32_LogicalDisk Where DeviceID = '" & strLogFileDrive & "'") 
                For Each objDisk in colDisks 
                    intFreeSpace = objDisk.FreeSpace 
                    intTotalSpace = objDisk.Size 
                    pctFreeSpace = intFreeSpace / intTotalSpace 
                    if (pctFreeSpace < ".50") then 
                        alertMessage = alertMessage & "Storage Group " _ 
                            & strStorageGroup & "'s TLOG Drive " _ 
                            & strLogFileDrive & " (on " & strComputer _ 
                            & ") has only " & FormatPercent(pctFreeSpace) _ 
                            & " disk free." & VbCrLf 
                    end if 
                Next 
                objRecordSet.MoveNext 
            Loop 
             
             
            '---Check C: and D: Drive Util--- 
            ' We only care about the local drives here, but you can add as many 
            ' as you want. 
            ' For some bizarre reason, I couldn't use the other method to  
            ' declare the array like I did when declaring the EX servers.  Weird. 
            arrDrives = Array("C:", "D:") 
             
            For each strLocalDrive in arrDrives 
            Set colDisks = objWMIService.ExecQuery _ 
                    ("Select * from Win32_LogicalDisk Where DeviceID = '" & strLocalDrive & "'") 
                For Each objDisk in colDisks 
                    intFreeSpace = objDisk.FreeSpace 
                    intTotalSpace = objDisk.Size 
                    pctFreeSpace = intFreeSpace / intTotalSpace 
                    if (pctFreeSpace < ".70") then 
                        alertMessage = alertMessage & "Local Disk " _ 
                            & strLocalDrive & " on " & strComputer _ 
                            & " has only " & FormatPercent(pctFreeSpace) _ 
                            & " disk free." & VbCrLf 
                    end if 
                Next 
            Next 
 
        '---Whoops, not pinging...--- 
    Else 
        alertMessage = alertMessage & strComputer & " is not pinging" & VbCrLf 
    End If 
'        i = i + 1 
'        WScript.Echo 
    'Next 
Next 
 
' Were any problems reported? 
if (len(alertMessage) > 0) then 
    alertMessage = "The following problems have been reported:" & VbCrLf _ 
        & VbCrLf & alertMessage 
'    wscript.echo alertMessage        'testing message 
    fnEmail() 
else 
'    WScript.Echo "No issues"        'testing message 
end if 
 
 
'=============================================================================== 
'==============================FUN WITH FUNCTIONS=============================== 
'=============================================================================== 
 
 
Function fnEmail() 
    Set objMessage        = CreateObject("CDO.Message") 
    objMessage.Subject    = "Exchange Monitoring has discovered issues" 
    objMessage.Sender    = "exchangeadministrator@domain.com" 
    objMessage.To        = "someonewhocares@domain.com; someoneelse@domain.com; exchangepager@domain.com" 
    objMessage.TextBody    = alertMessage 
 
 
    'This section provides the configuration information for the remote SMTP server. 
    objMessage.Configuration.Fields.Item _ 
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
 
    'Name or IP of Remote SMTP Server 
    objMessage.Configuration.Fields.Item _ 
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.subdomain.domain.com" 
 
    'Server port (typically 25) 
    objMessage.Configuration.Fields.Item _ 
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
    objMessage.Configuration.Fields.Update 
 
    objMessage.Send
End Function 
 
 
'====================================== 
'====================================== 
 
 
'Function fnPing(strComputer) 
Function IsAlive(strComputer) 
    ' by Phil Gordemer of ARH Associates 
    ' from http://www.tek-tips.com/viewthread.cfm?qid=1279504&page=3 
     
    '--- Test to see if host or url alive through ping --- 
    ' Returns True if Host responds to ping 
    ' 
    ' Though there are other ways to ping a computer, Win2K, 
    ' XP and different versions of PING return different error 
    ' codes. So the only reliable way to see if the ping 
    ' was sucessful is to read the output of the ping 
    ' command and look for "TTL=" 
    ' 
    ' strHost is a hostname or IP 
     
    const OpenAsASCII = 0 
    const FailIfNotExist = 0 
    const ForReading =  1 
    Dim objShell, objFSO, sTempFile, fFile 
    Set objShell = CreateObject("WScript.Shell") 
    Set objFSO = CreateObject("Scripting.FileSystemObject") 
    sTempFile = objFSO.GetSpecialFolder(2).ShortPath & "\" & objFSO.GetTempName 
     
    objShell.Run "%comspec% /c ping.exe -n 2 -w 500 " & strComputer & ">" & sTempFile, 0 , True 
    Set fFile = objFSO.OpenTextFile(sTempFile, ForReading, FailIfNotExist, OpenAsASCII) 
    Select Case InStr(fFile.ReadAll, "TTL=") 
        Case 0 
            IsAlive = False 
        Case Else 
            IsAlive = True 
    End Select 
    fFile.Close 
    objFSO.DeleteFile(sTempFile) 
    Set objFSO = Nothing 
    Set objShell = Nothing 
End Function 
 
 
'====================================== 
'======================================
