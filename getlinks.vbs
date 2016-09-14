' getlinks.vbs
' Copyright (C) 2016 lea2000
' 
' This program is free software; you can redistribute it
' and/or modify it under the terms of the GNU Lesser General
' Public License as published by the Free Software Foundation;
' either version 3 of the License, or (at your option) any
' later version.
' 
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
' Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public
' License along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

' On Error Resume Next

boolDeleteCache = False

'

Set objShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Set objStdOut = objFSO.GetStandardStream(1)
Set objStdErr = objFSO.GetStandardStream(2)

'

strScriptsFolder = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
strWindowsUpdateLog = objShell.ExpandEnvironmentStrings("%SYSTEMROOT%") & "\WindowsUpdate.log"

'

ReDim arrClassificationGUIDs(11)
arrClassificationGUIDs(0)  = "5c9376ab-8ce6-464a-b136-22113dd69801"
arrClassificationGUIDs(1)  = "434de588-ed14-48f5-8eed-a15e09a991f6"
arrClassificationGUIDs(2)  = "e6cf1350-c01b-414d-a61f-263d14d133b4"
arrClassificationGUIDs(3)  = "e0789628-ce08-4437-be74-2495b842f43b"
arrClassificationGUIDs(4)  = "e140075d-8433-45c3-ad87-e72345b36078"
arrClassificationGUIDs(5)  = "b54e7d24-7add-428f-8b75-90a396fa584f"
arrClassificationGUIDs(6)  = "9511d615-35b2-47bb-927f-f73d8e9260bb"
arrClassificationGUIDs(7)  = "0fa1201d-4330-4fa8-8ae9-b877473b6441"
arrClassificationGUIDs(8)  = "68c5b0a3-d1a6-4553-ae49-01d3a7827828"
arrClassificationGUIDs(9)  = "b4832bd8-e735-4761-8daf-37f882276dab"
arrClassificationGUIDs(10) = "28bc880e-0592-4cbf-8f95-c79b17911d5f"
arrClassificationGUIDs(11) = "cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83"

ReDim arrClassificationTypes(11)
arrClassificationTypes(0)  = "Application"
arrClassificationTypes(1)  = "Connectors"
arrClassificationTypes(2)  = "CriticalUpdates"
arrClassificationTypes(3)  = "DefinitionUpdates"
arrClassificationTypes(4)  = "DeveloperKits"
arrClassificationTypes(5)  = "FeaturePacks"
arrClassificationTypes(6)  = "Guidance"
arrClassificationTypes(7)  = "SecurityUpdates"
arrClassificationTypes(8)  = "ServicePacks"
arrClassificationTypes(9)  = "Tools"
arrClassificationTypes(10) = "UpdateRollups"
arrClassificationTypes(11) = "Updates"	

'

Set objRegExpURL = New RegExp
objRegExpURL.IgnoreCase = True
objRegExpURL.Global = True
objRegExpURL.Pattern = "([A-Za-z]{3,9})://([-;:&=\+\$,\w]+@{1})?([-A-Za-z0-9\.]+)+:?(\d+)?((/[-\+~%/\.\w]+)?\??([-\+=&;%@\.\w]+)?#?([\w]+)?)?"

Set objRegExpKB = New RegExp
objRegExpKB.IgnoreCase = True
objRegExpKB.Global = True
objRegExpKB.Pattern = "kb ?\d+"

'

If boolDeleteCache = True Then

	Set colRunningServices = objWMIService.ExecQuery ("SELECT StartMode, State FROM Win32_Service WHERE Name = 'wuauserv'")

	For Each objItem in colRunningServices
		If objItem.State = "Running" Then
			Wscript.Echo "Stopping wuauserv"
			objItem.StopService()
			Wscript.Sleep 3000
			Exit For
		End If	
	Next

	strSoftwareDistribution = objShell.ExpandEnvironmentStrings("%SYSTEMROOT%") & "\SoftwareDistribution"
	If objFSO.FolderExists(strSoftwareDistribution) Then
		Wscript.Echo "Deleting folder " & strSoftwareDistribution
		objShell.Exec "cmd /c rmdir /s /q """ & strSoftwareDistribution & """"
		Wscript.Sleep 3000
		If objFSO.FolderExists(strWindowsUpdateLog) Then
			Wscript.Echo "Couldn't delete " & strSoftwareDistribution
			Wscript.Quit
		End If
	End If

	If objFSO.FileExists(strWindowsUpdateLog) Then
		Wscript.Echo "Deleting file " & strWindowsUpdateLog
		objShell.Exec "cmd /c del /f /q """ & strWindowsUpdateLog & """"
		Wscript.Sleep 3000
		If objFSO.FileExists(strWindowsUpdateLog) Then
			Wscript.Echo "Couldn't delete " & strWindowsUpdateLog
			Wscript.Quit
		End If
	End If

	Set colRunningServices = objWMIService.ExecQuery ("SELECT StartMode, State FROM Win32_Service WHERE Name = 'wuauserv'")
	For Each objItem in colRunningServices
		If objItem.State = "Stopped" Then
			Wscript.Echo "Starting wuauserv"
			objItem.StartService()
			Wscript.Sleep 3000
			Exit For
		End If	
	Next

End If

'

strArch = "x86"
If objShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%") = "AMD64" Then strArch = "x64"

strVersion = ""
strProductType = "Desktop"
Set colItems = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
	If objItem.ProductType > 1 Then strProductType = "Server"
	strVersion = objItem.Version
Next

strPatchesFolder = strScriptsFolder & "msu\" & strArch & "\" & strProductType & "\" & strVersion

'If objFSO.FolderExists(strPatchesFolder) Then
'	objStdOut.WriteLine "Deleting folder " & strPatchesFolder
'	objShell.Exec "cmd /c rmdir /s /q " & strPatchesFolder
'	Wscript.Sleep 300
'End If

'

Set objSearcher = CreateObject("Microsoft.Update.Searcher")

objStdOut.Write "Getting not installed updates... "
Set objSearchResults = objSearcher.Search("IsInstalled=0")
Set colNotInstalledUpdates = objSearchResults.Updates
intNotInstalledCount = colNotInstalledUpdates.Count - 1
objStdOut.WriteLine intNotInstalledCount

'

strURLs = ""
urlDelemiter = ";"

If Not objFSO.FileExists(strWindowsUpdateLog) Then 
	objStdErr.WriteLine "Couldnt find " & strWindowsUpdateLog
	Wscript.Quit
End If

objStdOut.Write "Getting msu urls from WindowsUpdate.log... "
Set objWindowsUpdateLog = objFSO.OpenTextFile(strWindowsUpdateLog)
Do Until objWindowsUpdateLog.AtEndOfStream
	strLine = objWindowsUpdateLog.ReadLine
	If InStr(strLine, "http") > 0 And InStr(strLine, ".msu") > 0 Then
		Set objMatches = objRegExpURL.Execute(strLine)
		For Each objMatch in objMatches			
			strURL = objMatch.Value
			If InStr(LCase(strURLs), LCase(strURL)) = 0 Then strURLs = strURLs & strURL & urlDelemiter
		Next
	End If
Loop
objWindowsUpdateLog.Close

arrURLS = Split(strURLs, urlDelemiter)
intUrlCount = UBound(arrURLS) - 1
objStdOut.WriteLine intUrlCount

'

objStdOut.Write "Parsing not installed updates..."

For i = 0 to intNotInstalledCount
	
	Set objUpdate = colNotInstalledUpdates.Item(i)
	
	strTitle = objUpdate.Title	
	
	strKB = ""
	For Each strArticleID in objUpdate.KBArticleIDs
        strKB = "KB" & strArticleID
		Exit For
    Next	
	If Len(strKB) = 0 Then	
		Set objMatches = objRegExpKB.Execute(strTitle)
		For Each objMatch in objMatches
			strKB = UCase(objMatch.Value)
		Next	
	End If
	
	'
		
	strLastDeploymentChangeTime = objUpdate.LastDeploymentChangeTime	
	objDate = CDate(strLastDeploymentChangeTime)
	strSortableDate = Year(objDate) & Right("0" & Month(objDate), 2)
		
	'
	
	If Len(strKB) > 0 Then

		strKBURL = ""
		For j = 0 to intUrlCount
		
			strURL = arrURLS(j)
			If InStr(strURL, LCase(strKB)) > 0 Then
				strKBURL = strURL
				Exit For
			End If			
		
		Next
	
	End If
	
	'
	
	boolAutoSelect = objUpdate.AutoSelectOnWebSites
	
	strClassification = "999_NoCategory"

	Set objCategories = objUpdate.Categories

	boolFoundClassification = False
	For Each objCategory in objCategories
			
		strCategoryID = objCategory.CategoryID

		For k = 0 To UBound(arrClassificationGUIDs)
			If strCategoryID = arrClassificationGUIDs(k) Then
									
				strAutoSelect = k & "1"
				If boolAutoSelect = True Then strAutoSelect = k & "0"
				If Len(strAutoSelect) = 2 Then strAutoSelect = "0" & strAutoSelect
				
				strClassification = strAutoSelect & "_" & strSortableDate & "_" & arrClassificationTypes(k)
				boolFoundClassification = True
				Exit For
				
			End If
		Next
		
		If boolFoundClassification = True Then Exit For			

	Next
	
	'
	
	strCategoryFolder = strPatchesFolder & "\" & strClassification
	If Not objFSO.FolderExists(strCategoryFolder) Then
		objShell.Exec "cmd /c md """ & strCategoryFolder & """"
		Wscript.Sleep 300
		If Not objFSO.FolderExists(strCategoryFolder) Then
			objStdErr.WriteLine ""
			objStdErr.WriteLine "Couldn't create folder " & strCategoryFolder
			objStdErr.WriteLine ""
			Wscript.Quit
		End If
	End If
	
	'
	

	strLinks = strCategoryFolder & "\" & "links.txt"	
	If Len(strKBURL) = 0 Then
		If Len(strKB) = 0 Then
			strLinks = strCategoryFolder & "\" & "nokb.txt"
			strKBURL = strTitle
			objStdOut.WriteLine ""
			objStdOut.WriteLine  "Found no KB for " & strTitle
			objStdOut.WriteLine ""
		Else
			strLinks = strCategoryFolder & "\" & "nourl.txt"
			strKBURL = strKB
			objStdOut.WriteLine ""
			objStdOut.WriteLine  "Found no URL for " & strTitle
			objStdOut.WriteLine ""
		End If
	End If
	
	
	intMode = 2
	boolAlreadyThere = False
	
	If objFSO.FileExists(strLinks) Then 
		intMode = 8
		
		Set objLinksFile = objFSO.OpenTextFile(strLinks)
		Do Until objLinksFile.AtEndOfStream
			strLine = objLinksFile.ReadLine
			If InStr(strLine, strKBURL) > 0 Then 
				boolAlreadyThere = True
				objStdOut.WriteLine ""
				objStdOut.WriteLine  strTitel & "(" & strKB & ") already exists in " & strLinks
				objStdOut.WriteLine ""
				Exit Do
			End If
		Loop
		objLinksFile.Close
		
		
	End If
	
	If boolAlreadyThere = False Then
		
		Set objFile = objFSO.OpenTextFile (strLinks, intMode, True)
		objFile.WriteLine(strKBURL)
		objFile.Close
		
	End If
Next