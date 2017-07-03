' Created by :Tom Margrave  At Orasi Support
' File created:Wed Apr 14 2017
' Modified June 20 2017'
' File Name  UFTLog.vbs
'  VBScript UFTLogs.vbs is used to assist in setting HPE Unified Functional Testing (UFT),
'   UFT License, UFT API, and other logs'
' based on HPE document title : How to enable Unified Functional Testing (UFT) logs?
' Document ID : KM00467327 dated 2017-Apr-18

' TODO  consider maxSizeRollBackups'
bCore = True

If IsProcessRunning("UFT.exe") Then
    sTitle = "UFT is running  and needs to be stopped" & vbCrLf & "Stopping script."
    MsgBox sTitle, vbOKOnly  + vbCritical, "ERROR Process running"
    WScript.Quit
End If

pInstallLoc = getInstallLocation("HP Unified Functional Testing")

'starting at UFT 14 name changed for UFT'
If (Len(pInstallLoc) < 2) Then
    pInstallLoc = getInstallLocation("HPE Unified Functional Testing")
End If

iResponse = MsgBox("Do you want to reset all logs back to normal?", vbYesNo, "UFT Logs Reset")
Select Case iResponse
    Case VBYes
        bAsk = False
    Case vbNo
        bAsk = True
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'ask for core mode'
iResponse = MsgBox("Do you want All UFT log files deleted?", vbYesNo, "UFT Log delete")


Select Case iResponse
    Case VBYes
        deleteUFTLogs()
    Case vbNo
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

fileXmlCore = pInstallLoc & "bin\log.config.xml"

CreateBackup(fileXmlCore)


'ask for License mode'
If bAsk Then
    iResponse = MsgBox("Do you want UFT License in Debug mode?", vbYesNoCancel, "UFT License Logging options")
Else
    iResponse = vbNo
End If

Select Case iResponse
    Case VBYes
        licenseMode = "DEBUG"
        bCore = False
    Case vbNo 'Back to orignal'
        licenseMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'ask for LogCatPackMobileMode  Mobile mode'
If bAsk Then
    iResponse = MsgBox("Do you want UFT Mobile in Debug mode?", vbYesNoCancel, "UFT Mobile options")
Else
    iResponse = vbNo
End If

Select Case iResponse
    Case VBYes
        LogCatPackMobileMode = "DEBUG"
        bCore = False
    Case vbNo
        LogCatPackMobileMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'ask for APIMode  Mobile mode'
If bAsk Then
    iResponse = MsgBox("Do you want UFT API in Debug mode?", vbYesNoCancel, "UFT API options")
Else
    iResponse = vbNo
End If

Select Case iResponse
    Case VBYes
        APIMode = "DEBUG"
        bCore = False
    Case vbNo
        APIMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

' ask for RemoteAgent  '
If bAsk Then
    iResponse = MsgBox("Do you want UFT Remote Agent in Debug mode?", vbYesNoCancel, "UFT Remote Agent options")
Else
    iResponse = vbNo
End If

Select Case iResponse
    Case VBYes
        RAMode = "DEBUG"
        bCore = False
    Case vbNo
        RAMode = "OFF"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

'  ask for AutomationAgent
If bAsk Then
    iResponse = MsgBox("Do you want UFT Automation Agent (AOM) in Debug mode?", vbYesNoCancel, "UFT Automation Agent (AOM) options")
Else
    iResponse = vbNo
End If

Select Case iResponse
    Case VBYes
        AOMMode = "DEBUG"
        bCore = False
    Case vbNo
        AOMMode = "ERROR"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select

If bCore Then
    'ask for core mode'
    If bAsk Then
        iResponse = MsgBox("Do you want UFT Core in Debug mode?", vbYesNoCancel, "UFT Logging options")
    Else
        iResponse = vbNo
    End If
Else
    iResponse = VBYes
End If

Select Case iResponse
    Case VBYes
        coreMode = "DEBUG"
        If len(coreAR) < 2 Then
            coreAR = "RollingFileAppender"
        End If
    Case vbNo
        coreMode = "ERROR"
        coreAR = "BufferingForwardingAppender"
    Case Else
        myEcho("Quitting script")
        WScript.Quit
End Select


SetUFTCore(fileXmlCore)

'#####################'
'Set UFT Remote agent'
fileXml = pInstallLoc & "bin\log.config.RemoteAgent.xml"
SetUFTCore(fileXml)
SetXMLvalue fileXML, ".//logger[@name = 'LogCatRmtAgent']/level", RAMode


'#####################'
' Set UFT AOM'
MyFile = pInstallLoc & "bin\log.config.AutomationAgent.xml"
CreateBackup(MyFile)
SetUFTCore(MyFile)

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
    strComputer = "."
    Dim Process, strObject
    IsProcessRunning = False
    strObject = "winmgmts://" & strComputer
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
    If Not(supressNotes = 1) Then
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
        If installer.ProductInfo(productCode, "ProductName") = productName Then
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
    appDataLocation = objShell.ExpandEnvironmentStrings("%APPDATA%")
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
    Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel = impersonate}!\\" & strComputer & "\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & pName & "'")
    For Each objProcess in colProcessList
        objProcess.Terminate()
    Next
    Set colProcessList = Nothing
    Set objWMIService = Nothing
End Function

'**********************************************************************
' Sub Name: CreateBackup
' Purpose:  Create backup of xml file if does not exist
' Author: Tom Margrave
' Input:
'	fileXML
' Return: None
' Prerequisites: None
'**********************************************************************
Function CreateBackup(fileXML)
    'Check to see if file exist.'
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
End Function

 '**********************************************************************
 ' Function Name: SetUFTCore
 ' Purpose:  Set Corr files setting
 ' Author: Tom Margrave
 ' Input: file to be changed
 ' Return: None
 ' Prerequisites:
 '**********************************************************************
Function SetUFTCore(fileXml)
    '#####################'
    ' Set UFT core settings '
    '  File located at <UFT>\bin\log.config.xml'
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")

    xmlDoc.Async = "False"
    xmlDoc.load( fileXml)

    'set core logging
    SetXMLNodeValue xmlDoc,  "//log4net/root/level", coreMode
    SetXMLNodeValue xmlDoc,  "//log4net/root/appender-ref", coreAR

    ' Set License logging'
    SetXMLNodeValue xmlDoc,  ".//logger[@name = 'HP.UFT.License']/level", licenseMode

    'Set Mobile logging
    SetXMLNodeValue xmlDoc,  ".//logger[@name = 'LogCatPackMobile']/level", LogCatPackMobileMode

    '######### API Logs'``
    ' Checking to see if API debug already set'
    Set nNode = xmlDoc.selectsinglenode("//log4net/logger[@name = 'HP.ST']")

    If (APIMode = "DEBUG")  AND (nNode is Nothing) Then

        Set objRoot = xmlDoc.selectsinglenode("//log4net")
        Set objRecord  = xmlDoc.createElement("logger")
        objRecord.SetAttribute("name") = "HP.ST"

        objRoot.appendChild objRecord

        Set nNode = xmlDoc.selectsinglenode("//log4net/logger[@name = 'HP.ST']")
        Set objRecord  = xmlDoc.createElement("priority")
        objRecord.SetAttribute("value") = "ALL"
        nNode.appendChild objRecord

        Set objRecord  = xmlDoc.createElement("level")
        objRecord.SetAttribute("value") = "ALL"
        nNode.appendChild objRecord

        Set objRecord  = xmlDoc.createElement("appender-ref")
        objRecord.SetAttribute("ref") = "RollingFileAppender"
        nNode.appendChild objRecord

        Set objRecord  = xmlDoc.createElement("appender-ref")
        objRecord.SetAttribute("ref") = "ColoredConsoleAppender"
        nNode.appendChild objRecord
        '
        Set objRecord  = xmlDoc.createElement("appender-ref")
        objRecord.SetAttribute("ref") = "Recorder"
        nNode.appendChild objRecord
        '
        Set objRecord  = xmlDoc.createElement("appender-ref")
        objRecord.SetAttribute("ref") = "DebugAppender"
        nNode.appendChild objRecord
        '
        Set objRecord  = xmlDoc.createElement("appender-ref")
        objRecord.SetAttribute("ref") = "FileAppender"
        nNode.appendChild objRecord
    Else
        'Remove logging if set.
        Set nNodes = xmlDoc.selectNodes("//log4net/logger[@name = 'HP.ST']")
        For Each node In nNodes
            node.parentNode.removeChild(node)
        Next
    End If

    'Saving xmldoc
    ' On Error Resume Next
    '
    xmldoc.save(fileXml)
    If Err Then
        MsgBox "Error " & Err.Number & vbCrLf & Err.description & vbCrLf & "with file: " & fileXml, vbOKOnly  + vbCritical, "ERROR writing"
        WScript.Quit 1
    End If
    ' On Error GoTo 0

    Set queryNode = nothing
    set objRecord = nothing
    Set nNodes = nothing
    Set xmlDoc = nothing

End Function

 '**********************************************************************
 ' Function Name: SetXMLvalue
 ' Purpose: Set Element value
 ' Author: Tom Margrave
 ' Input:
 '      fileXML file with the xml
 '      myQuery the query to the elment
 '      myValue value to change
 ' Return: None
 ' Prerequisites:
 '**********************************************************************
Function SetXMLvalue(fileXML, myQuery, myValue)
    CreateBackup(fileXML)

    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    xmlDoc.Async = "False"
    xmlDoc.load fileXml

    ' Set Remote logging'
    Set queryNode = xmlDoc.selectSingleNode(myQuery)
    If Not(queryNode is Nothing) Then
        queryNode.Attributes.getNamedItem("value").Text = licenseMode
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

    Set queryNode = nothing
    set objRecord = nothing
    Set nNodes = nothing
    Set xmlDoc = nothing
End Function


 '**********************************************************************
 ' Function Name: SetXMLNodeValue
 ' Purpose: Set element value with node provied
 ' Author: Tom Margrave
 ' Input:
'       xmlDoc XML Document
'       xElement Element to change
'       xValue Valuse to change
 ' Return: None
 ' Prerequisites:
 '**********************************************************************
Function  SetXMLNodeValue( xmlDoc, xElement, xValue)
    Set queryNode = xmlDoc.selectSingleNode(xElement)

    If NOT((queryNode is Nothing) ) Then
        If (InStr(xElement,"appender-ref" ) > 2 )Then
            queryNode.Attributes.getNamedItem("ref").Text = xValue
        Else
        queryNode.Attributes.getNamedItem("value").Text = xValue
        End If
    End If
    'body
End Function
