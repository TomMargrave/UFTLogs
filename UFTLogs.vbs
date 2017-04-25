' Created by :Tom Margrave  At Orasi Support
' File created:Wed Apr 14 2017
' File Name  UFTLog.vbs
'  VBScript UFTLogs.vbs is used to assist in setting HPE Unified Functional Testing (UFT),
'   UFT License, UFT API, and other logs'
' based on HPE document title : How to enable Unified Functional Testing (UFT) logs?
' Document ID : KM00467327

' TODO  consider maxSizeRollBackups'

If IsProcessRunning("UFT.exe") Then
    sTitle = "UFT is running  and needs to be stopped" & vbCrLf & "Stopping script."
    MsgBox sTitle, vbOKOnly  + vbCritical, "ERROR Process running"
    WScript.Quit
End If

pInstallLoc = getInstallLocation("HP Unified Functional Testing")

'starting at UFT 14 name changed for UFT'
If Len(pInstallLoc < 2) Then
    pInstallLoc = getInstallLocation("HPE Unified Functional Testing")
End If

'ask for core mode'
iResponse = MsgBox("Do you want UFT log files deleted?", vbYesNoCancel, "UFT Log delete")

Select Case iResponse
    Case VBYes
        deleteUFTLogs()
    Case vbNo
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select


fileXml = pInstallLoc & "bin\log.config.xml"

If NOT(doesFileExist(fileXml)) Then
    sTitle = "Cannot find the xml to change. " & vbCrLf & fileXml & vbCrLf & "Exiting script"
    MsgBox sTitle, vbOKOnly  + vbCritical, "ERROR locating "
    WScript.Quit
End If

'Create back up '
If NOT(doesFileExist(fileXml & ".BAK")) Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    '
    objFSO.CopyFile fileXml, fileXml & ".BAK"
    If Err Then
        sTitle = "Error " & Err.Number & vbCrLf & Err.description & vbCrLf & "with files: " & fileXml & vbCrLf & fileXml & ".BAK"
        MsgBox sTitle, vbOKOnly  + vbCritical, "ERROR writing"
        WScript.Quit 1
    End If
    On Error GoTo 0

    Set objFSO = nothing
End If

'ask for core mode'
iResponse = MsgBox("Do you want UFT Core in Debug mode?", vbYesNoCancel, "UFT Logging options")

Select Case iResponse
    Case VBYes
        coreMode = "DEBUG"
        coreAR = "RollingFileAppender"
    Case vbNo
        coreMode = "ERROR"
        coreAR = "BufferingForwardingAppender"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select


'ask for License mode'
iResponse = MsgBox("Do you want UFT License in Debug mode?", vbYesNoCancel, "UFT License Logging options")

Select Case iResponse
    Case VBYes
        licenseMode = "DEBUG"
    Case vbNo
        licenseMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'ask for Help mode'
iResponse = MsgBox("Do you want UFT Help in Debug mode?", vbYesNoCancel, "UFT Help Logging options")

Select Case iResponse
    Case VBYes
        helpMode = "DEBUG"
    Case vbNo
        helpMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select


'ask for HTMLReportLoggerMode mode'
iResponse = MsgBox("Do you want UFT HTML Report Logger in Debug mode?", vbYesNoCancel, "UFT HTML Report Logging options")

Select Case iResponse
    Case VBYes
        HTMLReportLoggerMode = "DEBUG"
    Case vbNo
        HTMLReportLoggerMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'ask for LogCatPackMobileMode  Mobile mode'
iResponse = MsgBox("Do you want UFT Mobile in Debug mode?", vbYesNoCancel, "UFT Mobile options")

Select Case iResponse
    Case VBYes
        LogCatPackMobileMode = "DEBUG"
    Case vbNo
        LogCatPackMobileMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'ask for LogCatPackMobileMode  Mobile mode'
iResponse = MsgBox("Do you want UFT API in Debug mode?", vbYesNoCancel, "UFT API options")

Select Case iResponse
    Case VBYes
        APIMode = "DEBUG"
    Case vbNo
        APIMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'  File located at <UFT>\bin\log.config.xml'
Set xmlDoc = CreateObject("Microsoft.XMLDOM")

xmlDoc.Async = "False"
xmlDoc.load fileXml

'set core logging
Set nNode = xmlDoc.selectsinglenode ("//log4net/root/level")
nNode.Attributes.getNamedItem("value").Text = coreMode
Set nNode = xmlDoc.selectsinglenode ("//log4net/root/appender-ref")
nNode.Attributes.getNamedItem("ref").Text = coreAR

' Set License logging'
Set queryNode = xmlDoc.selectSingleNode(".//logger[@name = 'HP.UFT.License']/level")
queryNode.Attributes.getNamedItem("value").Text = licenseMode

'Set Help Engine logging.
Set queryNode = xmlDoc.selectSingleNode(".//logger[@name = 'HP.HelpEngine']/level")
queryNode.Attributes.getNamedItem("value").Text = helpMode

'Set the HTML Report Logger '
Set queryNode = xmlDoc.selectSingleNode(".//logger[@name = 'HTMLReportLogger']/level")
queryNode.Attributes.getNamedItem("value").Text = HTMLReportLoggerMode

'Set Mobile logging
Set queryNode = xmlDoc.selectSingleNode(".//logger[@name = 'LogCatPackMobile']/level")
queryNode.Attributes.getNamedItem("value").Text = LogCatPackMobileMode

'######### API Logs'``
' Checking to see if API debug already set'
Set nNode = xmlDoc.selectsinglenode("//log4net/logger[@name = 'HP.ST']")

If (APIMode = "DEBUG")  AND (nNode is Nothing) Then

    Set objRoot = xmlDoc.documentElement
    Set objRecord  = xmlDoc.createElement("logger")
    objRecord.SetAttribute("name") = "HP.ST"

    objRoot.appendChild objRecord

    Set nNode = xmlDoc.selectsinglenode("//log4net/logger[@name = 'HP.ST']")
    Set objRecord  = xmlDoc.createElement("priority")
    objRecord.SetAttribute("value")= "ALL"
    nNode.appendChild objRecord

    Set objRecord  = xmlDoc.createElement("level")
    objRecord.SetAttribute("value")= "ALL"
    nNode.appendChild objRecord

    Set objRecord  = xmlDoc.createElement("appender-ref")
    objRecord.SetAttribute("ref")= "RollingFileAppender"
    nNode.appendChild objRecord

    Set objRecord  = xmlDoc.createElement("appender-ref")
    objRecord.SetAttribute("ref")= "ColoredConsoleAppender"
    nNode.appendChild objRecord
    '
    Set objRecord  = xmlDoc.createElement("appender-ref")
    objRecord.SetAttribute("ref")= "Recorder"
    nNode.appendChild objRecord
    '
    Set objRecord  = xmlDoc.createElement("appender-ref")
    objRecord.SetAttribute("ref")= "DebugAppender"
    nNode.appendChild objRecord
    '
    Set objRecord  = xmlDoc.createElement("appender-ref")
    objRecord.SetAttribute("ref")= "FileAppender"
    nNode.appendChild objRecord

Else
    'Remove logging if set.
    Set nNodes = xmlDoc.selectNodes("//log4net/logger[@name = 'HP.ST']")
    For Each node In nNodes
        node.parentNode.removeChild(node)
    Next
End If

'Saving xmldoc
On Error Resume Next
'
strResult = xmldoc.save(fileXml)
If Err Then
    MsgBox "Error " & Err.Number & vbCrLf & Err.description & vbCrLf & "with file: " & fileXml, vbOKOnly  + vbCritical, "ERROR writing"
    WScript.Quit 1
End If
On Error GoTo 0
myEcho("Completed setting the logging for UFT")


'**********************************************************************
' Sub Name: IsProcessRunning
' Purpose:  Check to see if Excel is running before starting
' Author: Tom Margrave
' Input:
'   strProcess  process name to look for.'
' Return: Boolean of the status of the search
' Prerequisites:
'**********************************************************************
Function IsProcessRunning( pName)
    strComputer="."
    Dim Process, strObject
    IsProcessRunning = False
    strObject   = "winmgmts://" & strComputer
    For Each Process in GetObject(strObject).InstancesOf("win32_process")
    If UCase(Process.name) = UCase(pName) Then
        IsProcessRunning = True
        Exit Function
    End If
    Next
End Function

'**********************************************************************
' Sub Name: myEcho
' Purpose:  Display messages depending on flag supressNotes
' Author: Tom Margrave
' Input:
'	supressNotes
' Return: None
' Prerequisites:
'**********************************************************************
Function myEcho(strTemp)
    If Not(supressNotes=1) Then
        WScript.Echo strTemp
    End If
End Function


'**********************************************************************
' Sub Name: getInstallLocation
' Purpose:  Returns location where product is installed
' Author: Tom Margrave
' Input:
'	productName Name of the product as it appears in Control Panel
' Return: Non
' Prerequisites:
'**********************************************************************
Function getInstallLocation(productName)
    Dim installer
    Set installer = CreateObject("WindowsInstaller.Installer")
    getInstallLocation = ""
    For Each productCode In installer.Products
        If installer.ProductInfo(productCode, "ProductName")= productName Then
            getInstallLocation = installer.ProductInfo(productCode, "InstallLocation")
            Exit For
        End If
    Next
    Set installer = nothing
End Function


'**********************************************************************
' Sub Name: doesFileExist
' Purpose:  Checks to see if file exists
' Author: Tom Margrave
' Input:
'	mFile   File path to the check
' Return: Boolean if file exist
' Prerequisites:
'**********************************************************************
Function doesFileExist(mFile)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(mFile) Then
        doesFileExist = True
    Else
        doesFileExist = False
    End If
    Set objFSO = nothing
End Function

'**********************************************************************
' Sub Name: deleteUFTLogs
' Purpose:  Deletes all of the UFT logs
' Author: Tom Margrave
' Input:
'	None
' Return: None
' Prerequisites:
'   Function IsProcessRunning
'   Function killProcess
'**********************************************************************
Function deleteUFTLogs()
    Set objShell = CreateObject( "WScript.Shell" )
    appDataLocation=objShell.ExpandEnvironmentStrings("%APPDATA%")
    logsFldr = appDataLocation & "\Hewlett-Packard\UFT\Logs\"
    Set objShell = Nothing

    pName = "UFTRemoteAgent.exe"
    If IsProcessRunning(pName) Then
        sTitle = "UFT Agent is running  and needs to be stopped" & vbCrLf & "Killing process."
        MsgBox sTitle, vbOKOnly  + vbCritical, "ERROR Process running"
        killProcess(pName)
    End If

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set regexNumber = New RegExp

    regexNumber.Global = True
    regexNumber.IgnoreCase = True
    regexNumber.Pattern = "^\d+$"

    ' Delete Existing Files
    On Error Resume Next
    For Each oFile In objFSO.GetFolder(logsFldr).Files

        'Get only files that start with HP.'
        sFile = oFile.Name
        If StrComp(UCase(Left(sFile, 3)), "HP.", vbTextCompare) = 0 Then
            'Get Extension to lower case'
            sFileExt = LCase(objFSO.GetExtensionName(oFile.Name))
            If sFileExt = "log" Then
                oFile.Delete
            ElseIf regexNumber.Test(sFileExt) Then
                'look at extension to be a number and delete if True
                oFile.Delete
            End If
        End If
    Next
    On Error GoTo 0
    Set regexNumber = Nothing
    Set objFSO = Nothing
End Function

'**********************************************************************
' Sub Name: killProcess
' Purpose:  Kills process but will not kill system process
' Author: Tom Margrave
' Input:
'	pName  Name of the process to be killed iexplore.exe
' Return: None
' Prerequisites: None
'**********************************************************************
Function killProcess(pName)
    Const strComputer = "."
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & pName & "'")
    For Each objProcess in colProcessList
        objProcess.Terminate()
    Next
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Function
