'=========================================================================================================================================
' NAME: AD-COIT 2.546 (Active Directory Computer Object Inventory Tool)
' AUTHOR: Nick, wmi.wmi@gmail.com,
' DATE  : 8/5/2013
'=========================================================================================================================================
' Usage: "cscript scriptname.vbs -h"
' Notes: Under Script Configuration change the LDAP path to reflect the path for your organization, as well as the path to the CSV file
' Other requirements: "dsquery" must be installed locally in order to determine FSMO roles; download 2003 Resource kit.
' =========================================================================================================================================
' ChangeLog: 
' 5/30/2007 - 2.546, Added time stamp in csv output, and added DSquery output to csv if specified 
' 5/21/2007 - 2.542, Added -h functionality 
' 5/10/2007 - 2.451, Completed .csv basic functionality
' 5/8/207 - 2.449, Minor revisions to the following functions in support of .csv output  - OperatingSystem, BiosQuery, ProcQuery, DaylightSavingsEffect, SavDate
' 5/6/2007 - 2.48, Added .csv output
' 5/1/2007 - 2.4, Added LocalAdminEnum function to grab members of the local admin group
' 2/24/2006 - 2.3, Added Installed Application list
' 2/18/2006 - 2.2., Modified FSMO role query with an array
' 2/17/2006 - 2.1, Added Domain role query
' 2/16/2006 - 2.0, Added Domain Role detection
' 12/6/2005 - 1.9, Converted existing code into functions, added additional inventory objects
' 11/30/2005 - 1.5, Added free drive space function
' 11/25/2005 - 1.1, Added Smantec Antivirus Definition date inventory
' 8/31/2005 - 1.0, Initial feature release
'=========================================================================================================================================

'=== Script Configuration ================================================================================================================
strLdapPath = "'LDAP://DC=domain,DC=local'"   				' Set to relflect your LDAP path 
Const resultFilePath = "C:\scripts\ADCOITOutput.csv"				' Results files
'=== End Script Configuration ============================================================================================================

' Script Version Info
Dim strScriptVersion
strScriptVersion = "2.546"

Dim DisplayHelp, DisplayFSMO, DisplayMain, OutputCSV
Dim rootDSE
Dim DistinguishedName
Dim strInstallDateClean
Dim objDef,objDate,objVer,objRev,objOutFile,objNDate,objToday, strRole 
Dim objFSO
Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
Dim objResults
Set objResults = objFSO.CreateTextFile(resultFilePath, True) 
Dim objShell
Set objShell = Wscript.CreateObject("WScript.Shell")
Const MegabyteConversion = 1048576 					' See DiskSpace Function
Const Warning_Threshold = 1000  					' See DiskSpace Function

'objResults.WriteLine "Hostname,OS,SP,Install_Date,Manufacturer,BIOS,Serial,CPU,Role,Current_User,Model,RAM,DaylightSavingsOn,TimeZone,SAV_Date,Local_Hard_Drives,SQL_Server_Status,Instaled_Application,LocalAdmins"

'=== Main Logic ================================================================================================================

getCommandLineArgs 							' Seperates out 



If DisplayHelp = True Then
ShowHelp
Else
End if

Do while DisplayHelp = False

	'wscript.echo strLdapPath

	If OutputCSV = True Then
		objResults.WriteLine now				
	Else
	End if
	
	If DisplayFSMO = True Then
		dsqueryFSMOroles
	Else
	End if


	If OutputCSV = True Then
		'Builds CSV header 
		objResults.WriteLine "Hostname,OS,SP,Install_Date,Manufacturer,BIOS,Serial,CPU,Role,Current_User,Model,RAM,DaylightSavingsOn,TimeZone,SAV_Date,Local_Hard_Drives,SQL_Server_Status,Instaled_Application,LocalAdmins"
	Else
	End If

If DisplayMain = True Then
	On error resume next
	Const ADS_SCOPE_SUBTREE = 2
	Set objConnection = CreateObject("ADODB.Connection")
	Set objCommand =   CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Open "Active Directory Provider"
	Set objCOmmand.ActiveConnection = objConnection
	objCommand.CommandText = "Select Name, Location from " & strLdapPath & " Where objectCategory='computer'"
	objCommand.Properties("Page Size") = 1000
	objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
	Set objRecordSet = objCommand.Execute
	objRecordSet.MoveFirst
	Do Until objRecordSet.EOF
	    strComputer = objRecordSet.Fields("Name").Value
	        strCommand = "%comspec% /c ping -n 3 -w 1000 " & strComputer & ""
	    Set objExecObject = objShell.Exec(strCommand)
	strText=""

Do While Not objExecObject.StdOut.AtEndOfStream
        strText = objExecObject.StdOut.ReadAll()
        If Instr(strText, "Reply") > 0 Then
            Set objWMIService = GetObject _
                ("winmgmts:\\" & strComputer & "\root\cimv2")
		If Err.Number > 0 then
		  strErrorSystems =  strComputer & ", " & strErrorSystems 
		else
		  BuildHeader					
	          OperatingSystem 
	          BiosQuery
		  ProcQuery(strComputer)
		  DayLightSavingsEffect(strComputer)
		  SavDate(strComputer)
		  Diskspace
                  ListInstalledApplications			
		  LocalAdminEnum(strComputer)
		  SQLServerServices(strComputer)
		  BuildFooter
		  If OutputCSV = True Then
		  	objResults.WriteLine strComputer & "," & strOs & "," & strServPack & "," & strInstallDateClean & "," & strManuClean & "," & strBiosNameClean & "," & strDellTagClean & "," & strProcessor & "," & strDomainRole & "," & strUsernameReport & "," & strModel & "," & strRAM & "," & strDayLIghtSavings & "," & strTimeZone & "," & objNDate & "," & strLogicalDiskReport & "," & strSQLServerStatus & "," & strInstalledApplicationReport & "," & strMemberReport
		  Else
		  End if
		End if
	    Err.Clear 						'flushes error code from previous loop
        Else
           UnavailableSystems =  strComputer & ", " & UnavailableSystems

	   End If
    Loop
    objRecordSet.MoveNext
Loop
Wscript.Echo "The following systems were unavailable: " & UnavailableSystems 
Wscript.Echo " "
Wscript.Echo "The following systems were on, but returned an error: " & strErrorSystems 
Else														
End if
DisplayHelp = True
Loop


Function ListInstalledApplications
strInstalledApplicationReport = ""
Set colItemsProduct = objWMIService.ExecQuery("Select * from Win32_Product")
 Wscript.echo "Workstation Application List:"
 Wscript.Echo "-----------------------------"
For Each objItemProduct in colItemsProduct
 Wscript.echo objItemProduct.Description &  " (v." & objItemProduct.Version & ")"
 'Wscript.echo "Location: " & objItem.InstallLocation & ", " 'uncomment this line to get path information
 strInstalledApplicationReport = strInstalledApplicationReport & objItemProduct.Description &  " (v." & objItemProduct.Version & ")" & "; "
Next

End Function

Function dsqueryFSMOroles
Dim objShell, ArrFSMOCmds(4)
Set objShell = CreateObject("WScript.Shell")

Cmd1 = "cmd /c"
Cmd2 = "dsquery server -hasfsmo "
ArrFSMOCmds(0) = "schema"
ArrFSMOCmds(1) = "rid"
ArrFSMOCmds(2) = "name"
ArrFSMOCmds(3) = "infr"
ArrFSMOCmds(4) = "PDC"

For Each ArrFSMOCmds in ArrFSMOCmds
Set objExec = objShell.Exec(Cmd1 & Cmd2  & ArrFSMOCmds)
		strExecResults = LCase(objExec.StdOut.ReadAll)
ArrFSMOOutput = ArrFSMOCmds & ": " & strExecResults
Wscript.Echo ArrFSMOCmds & ": " & strExecResults

If OutputCSV = True Then
	objResults.WriteLine ArrFSMOOutput			' Clean-up Output such that commas are handled
End If

Next

End Function

Function BuildHeader
Wscript.Echo "-------------------------------------------------------------" 'disabled for CSV output
Wscript.Echo "Computer Name: " & strComputer
Wscript.Echo "-------------------------------------------------------------" 'disabled for CSV output
End Function
Function OperatingSystem
   Set colItems = objWMIService.ExecQuery ("Select * From Win32_OperatingSystem")  
   For Each objItem in ColItems
     'strOs = objItem.Caption							' returns Operating System|c:\locationofpartition
     strOs = objItem.Name
     strOs = Split(strOs, "|")(0)						' seperates OS from partition and using the first string0 
     strServPack = objItem.CSDVersion
     strInstallDate = objItem.InstallDate
   Next
   Wscript.Echo "Operating System: " & strOs
      Wscript.Echo  "Service Pack: " & strServPack
      Wscript.Echo "Install Date: " & WMIDateStringToDate(strInstallDate)
      strInstallDateClean = WMIDateStringToDate(strInstallDate)
End Function

Function BiosQuery
	 Set colItems2 = objWMIService.ExecQuery("Select * from Win32_BIOS",,48)
         For Each objItem in colItems2
	   strDellTag = objItem.SerialNumber
	   strManu = objItem.Manufacturer
	   strBiosName = objitem.Name
	 Next
	 WScript.Echo "Manufacturer: " & strManu
	 Wscript.Echo "BIOS Name: " & strBiosName
	 Wscript.Echo "Service Tag: " & strDellTag
	 strManuClean = strManu
	 strBiosNameClean = strBiosName
	 strDellTagClean = strDellTag
End Function

Function WMIDateStringToDate(dtmDate)
WScript.Echo dtm: 
	WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
	Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
	& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
End Function
Function ProcQuery(shost)
Set colItems4 = objWMIService.ExecQuery("SELECT * FROM Win32_Processor")
 
 For Each objItem in colitems4
	strProcessor = objItem.Name
 Next     
 Wscript.Echo "Processor type: " & strProcessor
End Function

Function DayLightSavingsEffect(shost)
Set colItems3 = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)

For Each objItem in colitems3
	        strDomainRole = objItem.DomainRole
		strUserName = objItem.Username
		strModel = objItem.Model
		strRAM = objItem.TotalPhysicalMemory
		strTimeZone = (objItem.CurrentTimeZone / 60)
		strDayLightSavings = objItem.DaylightInEffect
	    Next

		Select Case strDomainRole 
        	Case 0
          	strRole = "Standalone workstation"
        	Case 1
          	strRole = "Member workstation"
        	Case 2
          	strRole = "Standalone server"
        	Case 3
          	strRole = "Member server"
        	Case 4
          	strRole = "Backup domain controller"
        	Case 5
          	strRole = "PDC Emulator"
      		End Select	    
	    Wscript.Echo "Domain Role: " & strRole
	    Wscript.Echo "Current User: " & strUsername
	    Wscript.Echo "Model: " & strModel
	    Wscript.Echo "RAM: " & strRAM
	    Wscript.Echo "Daylight Savings in effect: " & strDayLIghtSavings
	    Wscript.Echo "Time Zone: " & strTimeZone
	    strUsernameReport = strUsername
	    strUsername = ""
End Function

Function SavDate(shost)
strCmdRun = "cmd /c"
strRegQ = "reg query "
strRegKey = "\HKLM\SOFTWARE\Symantec\SharedDefs\"
strCmdSw = " /v "
strRegKey2 = "DEFWATCH_10"
		Set objExec = objShell.Exec(strCmdRun & strRegQ & "\\" & shost & strRegKey & strCmdSw & strRegKey2)
		strExecResults = LCase(objExec.StdOut.ReadAll)
		objVer = Right(strExecResults,16)
		objRev = Right(objVer,7)
		objDate = Left(objVer,8)
		objYear = Left(objDate,4)
		objMonth = Mid(objDate,5,2)
		objDay = Right(objDate,2)
		objNDate = CDATE(objMonth &"/"& objDay &"/"& objYear)
Wscript.Echo "Symantec Antivirus Definition Date: " & objNDate &" Rev. "& objRev & " "
End Function
Function DiskSpace
	strLogicalDiskReport = ""
	Set colLogicalDisk = objWMIService.InstancesOf("Win32_LogicalDisk")
	If colLogicalDisk.Count = 0 Then
		Wscript.Echo "No drives exist on this computer."
	Else
		For Each objLogicalDisk in colLogicalDisk
			If objLogicalDisk.DriveType = 3 Then
				FreeMegaBytes = objLogicalDisk.Freespace / MegabyteConversion
				FreeMegaBytes_String = Int(FreeMegaBytes) & " MB Free"
			   Wscript.Echo objLogicalDisk.DeviceID & ", " & FreeMegaBytes_String & ", " & objLogicalDisk.Description
			   strLogicalDiskReport = strLogicalDiskReport & FreeMegaBytes_String & " " & objLogicalDisk.Description & " "
			End if
		Next
	End if
End Function
Function BuildFooter
Wscript.Echo "-------------------------------------------------------------"
End Function

Function LocalAdminEnum(shost)
		strMemberReport = ""
		Set objLocalAdminGroup = GetObject("WinNT://" & shost & "/Administrators")
		Wscript.Echo shost & " local Administrator Group Membership: "
		For Each objMember In objLocalAdminGroup.Members
			WScript.Echo objMember.Name
			strMemberReport = strMemberReport & " " & objMember.Name
		Next
End Function


Function SQLServerServices(shost)
	
	Set colServices = objWMIService.ExecQuery ("Select * from Win32_Service Where Name = 'MSSQLServer'")
		strSQLServerStatus = ""
		If colServices.Count > 0 Then
		    For Each objService in colServices
		        Wscript.Echo "SQL Server is " & objService.State & "."
		        strSQLServerStatus = objService.State
		    Next
		Else
		    Wscript.Echo "SQL Server is not installed on this computer."
		    strSQLServerStatus = "Not Installed"
		End If

		Wscript.Echo strSQLServerStatus
End Function



Sub getCommandLineArgs()
Set objArgs = WScript.Arguments
If (objArgs.Count > 0) Then
	For nArgs = 0 To objArgs.Count - 1
		SetArgs objArgs(nArgs)
	Next
	Else
		WScript.Echo "For help: cscript.exe scriptname.vbs -h"
		
	End If
'DisplayHelp = False
End Sub

Sub ShowHelp
Wscript.Echo "AD-COIT: " & strScriptVersion
WScript.Echo "============================================================================="
WScript.Echo "Using: " & strLdapPath & ", " & resultFilePath
WScript.Echo "============================================================================="
WScript.Echo "Usage: 	cscript.exe scriptname.vbs [options] [output]"
WScript.Echo "Example:	cscript.exe scriptname.vbs -f -m -c"
WScript.Echo "		cscript.exe scriptname.vbs -m -c"
WScript.Echo "============================================================================="
Wscript.Echo "Set Options"
Wscript.Echo "-f	- Display FSMO roles using DSQUERY (must have DSQUERY installed)"
WScript.Echo "-m	- Execute Main set of logic - this does all of the inventory work"
WScript.Echo "-h	- Displays this Help page"
Wscript.Echo "Set Output"
WScript.Echo "-c	- Creates .CSV output file"
End Sub 

Sub SetArgs(strOption)
dim strFlag
strFlag = Left(strOption,2)
	Select Case strFlag
		Case "-h"
			DisplayHelp = True
		Case "-f"
			DisplayFSMO = True
		Case "-m"
			DisplayMain = True
		Case "-c"
			OutputCSV = True
	End Select
End Sub
