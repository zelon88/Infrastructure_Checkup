'File Name: Infrastructure_Checkup.vbs
'Version: v1.2, 11/25/2022
'Author: Justin Grimes, 3/5/2019

' --------------------------------------------------
'Declare explicit variables to be used during the session.
Option Explicit 
Dim fileSystem, windowsVersion, oShell, computerName, mailData, mailFile, message, logPath, strSafeDate, strSafeTime, _
 strDateTime, logfile, objLogFile, arguments, arg, args, error, appPath, verbose, email, logging, emailData, _
 logData, dataDir, computerDir, OutputBox, processArch, sysArch, WshProcEnv, messageData, archType, checkupResults, _
 requiredDir, computerDir0, computerDir1, argList, requiredDirs, requiredDirectories, mainError, dxDiagInfo, WshFinished, _
 WshError, taskInfo, msInfo, diskInfo, companyName, companyAbbr, companyDomain, toEmail
' --------------------------------------------------

' --------------------------------------------------
'Set variables values for the session.
Set oShell = WScript.CreateObject("WScript.Shell")
Set arguments = WScript.Arguments
Set WshProcEnv = oShell.Environment("Process")
Set fileSystem = CreateObject("Scripting.FileSystemObject") 
computerName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
strSafeDate = DatePart("yyyy",Date)&Right("0"&DatePart("m",Date), 2)&Right("0"&DatePart("d",Date), 2)
strSafeTime = Right("0"&Hour(Now), 2)&Right("0"&Minute(Now), 2)&Right("0"&Second(Now), 2)
strDateTime = strSafeDate&"-"&strSafeTime
  ' ----------
  ' Company Specific variables.
  ' Change the following variables to match the details of your organization.
  
  ' The "appPath" is the full absolute path for the script directory, with trailing slash.
  appPath = "\\server\Scripts\Infrastructure_Checkup\"
  ' The "logPath" is the full absolute path for where network-wide logs are stored.
  logPath = "\\server\Logs"
  ' The "dataDir" is where all generated reports will be stored. 
  'Subdirectories for computer name and report time will be created in this directory.
  dataDir = "\\server\Logs\Computers\"
  ' The "companyName" the the full, unabbreviated name of your organization.
  companyName = "The Company Inc."
  ' The "companyAbbr" is the abbreviated name of your organization.
  companyAbbr = "TCI"
  ' The "companyDomain" is the domain to use for sending emails. Generated report emails will appear
  ' to have been sent by "COMPUTERNAME@domain.com"
  companyDomain = "thecompany.com"
  ' The "toEmail" is a valid email address where notifications will be sent.
  toEmail = "helpdesk@thecompany.com"
  ' ----------
logfile = logPath&"\"&computerName&"-"&strDateTime&"-infrastructure_checkup.txt"
mailFile = appPath&"Warning.mail" 
computerDir0 = dataDir&computerName
computerDir1 = computerDir0&"\Checkups"
computerDir = computerDir1&"\"&strSafeDate
requiredDirectories = Array(appPath, dataDir, computerDir0, computerDir1, computerDir, logPath)
mainError = False
' --------------------------------------------------

' --------------------------------------------------
'Retrieve the specified arguments.
'Supported arguments are:
' -v  -  Verbose operation. Output any messages to a MsgBox.
' -e  -  Email operation. Output any messages to an email.
' -l  -  Log operation. Output any messages to a logfile.
' -s  -  Do not output any messages to a MsgBox.
'Returns an array of arguments in the order listed below.
Function ParseArgs(args)
  Dim outputArray(3)
  error = verbose = email = logging = False
  outputArray(0) = outputArray(1) = outputArray(2) = outputArray(3) = False
  For Each arg In args
    'Verbose
    If arg = "-v" Then
      outputArray(0) = True
    End If
    'Email
    If arg = "-e" Then
      outputArray(1) = True
    End If
    'Logging
    If arg = "-l" Then
      outputArray(2) = True
    End If
    'Silent mode (enable email+logging, no verbosity)
    If arg = "-s" Then
      outputArray(0) = False
      outputArray(1) = True
      outputArray(2) = True
      outputArray(3) = True
    End If
  Next
  If outputArray(0) = False And outputArray(1) = False And outputArray(2) = False And outputArray(3) = False Then
    error = True
    ParseArgs = False
  End If
  If error = -1 Then
    ParseArgs = outputArray
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
' Verify that all required directories exist.
Function VerifyInstallation(requiredDirs)
  error = VerifyInstallation = False
  If IsArray(requiredDirs) Then
    For Each requiredDir In requiredDirs
    On Error Resume Next
      'Verify the supplied directory.
      If Not fileSystem.FolderExists(requiredDir) Then
        fileSystem.CreateFolder(requiredDir)
        If Not fileSystem.FolderExists(requiredDir) Then
          error = True
        End If
      End If
    Next
    If error = -1 Then
      VerifyInstallation = True
    End If
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to send the notification email when -e is set.
'Returns "True" if email sent sucessfully, "False" on failure.
Function SendEmail(mailFile, mailContent) 
  error = True
  If fileSystem.FileExists(mailFile) Then
    fileSystem.DeleteFile(mailFile)
  End If
  If Not fileSystem.FileExists(mailFile) Then
    Set mailData = fileSystem.CreateTextFile(mailFile, True, True)
    mailData.Write mailContent
    mailData.Close
  End If
  If fileSystem.FileExists(mailFile) Then
    error = False
    oShell.exec appPath&"sendmail.exe "&mailFile
  End If
  SendEmail = error
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create a log file when -l is set.
'Returns "True" if logfile exists, "False" on error.
Function CreateLog(logFile, message)
  error = True
  If message <> "" Then
    Set objLogFile = fileSystem.CreateTextFile(logFile, True)
    objLogFile.WriteLine(message)
    objLogFile.Close
  End If
  If fileSystem.FileExists(logFile) Then
    error = False
  End If
  CreateLog = error
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to create a MsgBox when -v is set.
'Returns "True" if a MsgBox was displayed, "False" on error.
Function OutputMessage(message)
  error = True
  If message <> "" Then
    error = False
    OutputBox = MsgBox(message, 64, "Infrastructure Checkup")
  End If
  OutputMessage = error
End Function
' --------------------------------------------------

' --------------------------------------------------
'Determine the local system architecture so we don't encounter errors when we run tests.
'Returns either "x86" for 32 bit systems or "AMD64" for 64 bit systems.
Function DetermineArch()
  processArch = sysArch = DetermineArch = False
  processArch = WshProcEnv("PROCESSOR_ARCHITECTURE") 
  If processArch = "x86" Then    
    sysArch = WshProcEnv("PROCESSOR_ARCHITEW6432")
    If sysArch = ""  Then    
      sysArch = "x86"
    End if    
  Else    
    sysArch = processArch    
  End If
  If sysArch = "x86" Or sysArch = "AMD64" Then
    DetermineArch = sysArch
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to retrieve and store the Windows version.
'Uses the Windows "ver" command.
Function GetWindowsVersion(computerDir, strDateTime)
  Dim winVersionCacheFile, windowsVersion, winVersionCache, versionInfo
  windowsVersion = GetWindowsVersion = False
  versionInfo = ""
  winVersionCacheFile = computerDir&"\Windows_Version_"&strDateTime&".txt"
  Set windowsVersion = oShell.exec("wmic os get Caption,CSDVersion /value")
  Select Case windowsVersion.Status
    Case WshFinished
    versionInfo = Trim(windowsVersion.StdOut.ReadAll)
  End Select
  Set winVersionCache = fileSystem.CreateTextFile(winVersionCacheFile, True, False)
  winVersionCache.WriteLine(versionInfo)
  winVersionCache.Close
  If fileSystem.FileExists(winVersionCacheFile) Then
    GetWindowsVersion = versionInfo
  End If

End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to retrieve and store the DXDiag information.
'Uses the Windows "dxdiag" command.
Function GetDXInfo(archType, computerDir, strDateTime)
  Dim dxCacheFile, dxInfo, dxCache, extraQuery
  If archType = "AMD64" Then
    extraQuery = " /64bit" 
  End If
  dxInfo = GetDXInfo = False
  dxCacheFile = computerDir&"\DXDiag_"&strDateTime&".txt"
  Set dxInfo = oShell.exec("dxdiag"&extraQuery&" /t "&dxCacheFile)
  If fileSystem.FileExists(dxCacheFile) Then
    GetDXInfo = dxInfo
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to retrieve and store running task information.
'Uses the Windows "tasklist" command.
Function GetTaskInfo(computerDir, strDateTime)
  Dim taskCacheFile, taskInfo, taskCache
  taskInfo = GetTaskInfo = False
  taskCacheFile = computerDir&"\Running_Tasks_"&strDateTime&".txt"
  Set taskInfo = oShell.exec("tasklist /v")
  Select Case taskInfo.Status
    Case WshFinished
    taskInfo = Trim(taskInfo.StdOut.ReadAll)
  End Select
  Set taskCache = fileSystem.CreateTextFile(taskCacheFile, True, False)
  taskCache.WriteLine(taskInfo)
  taskCache.Close
  If fileSystem.FileExists(taskCacheFile) Then
    GetTaskInfo = taskInfo
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to retrieve and store msinfo32 information.
'Uses the Windows "msinfo32" command.
Function GetMSInfo(computerDir, strDateTime)
  Dim msCacheFile, msInfo, msCache
  msInfo = GetMSInfo = False
  msCacheFile = computerDir&"\MSInfo32_"&strDateTime&".txt"
  Set msInfo = oShell.exec("msinfo32 /report """&msCacheFile&"""")
  Select Case msInfo.Status
    Case WshFinished
    msInfo = Trim(msInfo.StdOut.ReadAll)
  End Select
  If fileSystem.FileExists(msCacheFile) Then
    GetMSInfo = msInfo
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'A function to retrieve and store disk information.
'Uses the Windows "df" command.
Function GetDiskInfo(computerDir, strDateTime)
  Dim diskCacheFile, diskInfo, diskCache
  diskInfo = GetDiskInfo = False
  diskCacheFile = computerDir&"\Disk_Info_"&strDateTime&".txt"
  Set diskInfo = oShell.exec("fsutil volume diskfree C:")
  Select Case diskInfo.Status
    Case WshFinished
    diskInfo = Trim(diskInfo.StdOut.ReadAll)
  End Select
  Set diskCache = fileSystem.CreateTextFile(diskCacheFile, True, False)
  diskCache.WriteLine(diskInfo)
  diskCache.Close
  If fileSystem.FileExists(diskCacheFile) Then
    GetDiskInfo = diskInfo
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'Perform checkup routine.
Function PerformCheckup(archType, verbose, email, logging, computerDir, strDateTime)
  If (fileSystem.FolderExists(computerDir)) Then 
    windowsVersion = GetWindowsVersion(computerDir, strDateTime)
    dxDiagInfo = GetDXInfo(archType, computerDir, strDateTime)
    taskInfo = GetTaskInfo(computerDir, strDateTime)
    msInfo = GetMSInfo(computerDir, strDateTime)
    diskInfo = GetDiskInfo(computerDir, strDateTime)
    'Check if all reports were genereated sucessfully.
    If vartype(windowsVersion) <> 8 Or vartype(dxDiagInfo) = 8 Or vartype(taskInfo) = 8 Or vartype(msInfo) = 8 Or vartype(diskInfo) = 8 Then
      mainError = True
    End If
  End If
End Function
' --------------------------------------------------

' --------------------------------------------------
'The main logic of the program.

'Start by parsing the supplied arguments.
argList = ParseArgs(arguments)
If IsArray(argList) Then
  verbose = argList(0)
  email = argList(1)
  logging = argList(2)
  'Verify that all required directories exist and the application is safe to run.
  If VerifyInstallation(requiredDirectories) = -1 Then
    archType = DetermineArch() 
    'Make sure we can determine the architecture of the target system before executing anything.
    If archType = "x86" Or archType = "AMD64" Then
      checkupResults = PerformCheckup(archType, verbose, email, logging, computerDir, strDateTime)
      If verbose Then
        messageData = "This is a notification from the " & companyName & " Network to inform you that the device "&computerName&" has completed a checkup. The results of the checkup are located in "&computerDir&"."& _
         vbNewLine&vbNewLine&"Please log-in and verify the generated reports for "&computerName&"."
        Call OutputMessage(messageData) 
      End If
      If email = -1 Then
        emailData = "To: " & toEmail & ""&vbNewLine&"From: "&computerName&"@" & companyDomain & ""&vbNewLine&"Subject: " & companyAbbr & " Infrastructure Checkup!!"&vbNewLine& _
         "This is an automatic email from the " & companyName & " Network to notify you that the device "&computerName&" has completed a checkup. The results of the checkup are located in "&computerDir&"."& _
         vbNewLine&vbNewLine&"Please log-in and verify the generated reports for "&computerName&"."&vbNewLine&vbNewLine& _
         "This check was generated by "&computerName&" and is performed as a monthly scheduled task."&vbNewLine&vbNewLine&"Script: ""Infrastructure_Checkup.vbs"""
        Call SendEmail(mailFile, emailData) 
      End If
      If logging = -1 Then
        logData = "This is an automatic message from the " & companyName & " Network to notify you that the device "&computerName&" has completed a checkup. The results of the checkup are located in "&computerDir&"."& _
         vbNewLine&vbNewLine&"Please log-in and verify the generated reports for "&computerName&"."&vbNewLine&vbNewLine& _
         "This check was generated by "&computerName&" and is performed as a monthly scheduled task."&vbNewLine&vbNewLine&"Script: ""Infrastructure_Checkup.vbs"""
        Call CreateLog(logFile, logData) 
      End If
    End If
  End If
End If
' --------------------------------------------------


