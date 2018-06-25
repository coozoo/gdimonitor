Option explicit
Dim strCurrentFolder,timestamp, starttimestamp,ArgObj,flag,headerstr,timeBetween,exeArj,processname,prevArgument,examplestring
headerstr="Timestamp,Process ID,Process Name,Pen,ExtPen,Brush,Bitmap,Font,Palette,Region,DC,Metafile DC,Enhanced Metafile DC,Other GDI,GDI Total,All GDI"
starttimestamp=DateTimeMS(0)
Set ArgObj = WScript.Arguments
timeBetween=0
processname=""

examplestring=vbCrLf +"   Example:" + vbCrLf+ "   "+ Wscript.ScriptName+ " -delay 10 -name notepad"

If WScript.Arguments.Count > 1 Then
prevArgument=""
dim i
	For i = 0 To WScript.Arguments.Count-1
		if (prevArgument="-name" OR prevArgument="-delay") AND (ArgObj(i)="-name" OR ArgObj(i)="-delay") then
			WScript.echo "-name and -delay expected input" + examplestring
			WScript.quit
		end if
		if prevArgument="-name" then
			processname=ArgObj(i)
		elseif prevArgument="-delay" then
			timeBetween=ArgObj(i)*1000
		end if
		prevArgument=ArgObj(i)
		if prevArgument="-delay" and i=WScript.Arguments.Count-1 then
			WScript.echo "Unknown or incorect arguments" + examplestring
			WScript.quit
		end if
	Next
Else
	WScript.echo "Incorrect Arguments "+ examplestring
	WScript.quit
end if
i=0

if timeBetween=0 then
	timeBetween=5000
end if
if processname="" then
	WScript.echo "Process name can't be empty"
	WScript.quit	
end if

If isSixtyFour=0 then
	exeArj="\GDIView.exe /scomma GDItemp"
else
	exeArj="\GDIView64.exe /scomma GDItemp"
end if	
flag=1
'######################################################################
'main loop
'######################################################################
Do while (flag=1)
	dumpGDIview
	FindProcessNodes
	WScript.Sleep timeBetween
loop

'######################################################################
' dump gdi into file
'######################################################################
Function dumpGDIview
	Const HIDDEN_WINDOW = 12
	dim BinPath
	Dim objResult, objShell
	Set objShell = WScript.CreateObject("WScript.Shell") 
	strCurrentFolder=objShell.CurrentDirectory
	BinPath = strCurrentFolder & exeArj
	objResult = objShell.Run(BinPath, 1, True)
	timestamp = DateTimeMS(1)
End Function

'######################################################################
' Find processname nodes inside gdidump
'######################################################################
Function FindProcessNodes()
	dim objRegEx,objFSO,objFile,strSearchString,colMatches,a,strMatch
	Const ForReading = 1
	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Pattern = processname
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strCurrentFolder & "\GDItemp", ForReading)
	Do Until objFile.AtEndOfStream
		strSearchString = objFile.ReadLine
		Set colMatches = objRegEx.Execute(strSearchString)  
		If colMatches.Count > 0 Then
			For Each strMatch in colMatches   
'				msgbox strSearchString
				a=Split(strSearchString,",")
'				msgbox a(0)
				writestring timestamp &","& strSearchString, strCurrentFolder & "\gdistat\" & a(0) & "GDIstat" & starttimestamp & ".csv"
			Next
		End If
	Loop
	objFile.Close
End Function

'######################################################################
' Write string into file 
' writestring("blabla", "log.csv")
'######################################################################
Function writestring(logstring, logfile)
	Dim logFSO, objOutFile, sWorkingFileName
	Const FOR_APPENDING = 8
	Set logFSO = CreateObject("Scripting.FileSystemObject")
	'sWorkingFileName = "\\servername\sharename\logfilename.txt"
	sWorkingFileName=logfile
	CreateFolderIfNo(strCurrentFolder & "\gdistat\")
	If logFSO.FileExists(sWorkingFileName) Then
		Set objOutFile = logFSO.OpenTextFile(sWorkingFileName, FOR_APPENDING)
	Else
		Set objOutFile = logFSO.CreateTextFile(sWorkingFileName)
		objOutFile.WriteLine headerstr
	End If
	objOutFile.WriteLine logstring
	objOutFile.close
End Function

'######################################################################
' date time in milisecconds precission
' DateTimeMS("1") - returns date time with sec and ms
' DateTimeMS("any number") - returns date time without sec and ms
'######################################################################
Function DateTimeMS(dformat)
dim sNow,sYear,sMonth,sDay,sHour,sMinute,sSecond,sMilliSecond,sFullDateTime
	sNow = Now
	sYear = Year(sNow)
	sMonth = Month(sNow)
	sDay = Day(sNow)
	sHour = Hour(sNow)
	sMinute = Minute(sNow)
	sSecond = Second(sNow)
	sMilliSecond = Timer * 1000 mod 1000
	if dformat =1 then
		sFullDateTime = sDay & "/" & sMonth & "/" & sYear & " " & sHour & ":" & sMinute &":" & sSecond & "." & sMilliSecond
	else
		sFullDateTime = sDay & "." & sMonth & "." & sYear & "-" & sHour & "-" & sMinute
	end if
	DateTimeMS=sFullDateTime
End Function

'######################################################################
' Create folder
' CreateFolderIfNo("foldername")
'######################################################################
Function CreateFolderIfNo(foldernamestr)
	dim filesys, newfolder
	set filesys=CreateObject("Scripting.FileSystemObject")
	If  Not filesys.FolderExists(foldernamestr) Then
		newfolder = filesys.CreateFolder (foldernamestr)
		'Response.Write "A new folder '" & newfolder & "' has been created"
	End If
	CreateFolderIfNo=newfolder
End Function

'######################################################################
' Check if 64
'
'######################################################################
Function isSixtyFour()
	Dim WshShell
	Dim OsType
	Set WshShell = CreateObject("WScript.Shell")
	OsType = WshShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
	If OsType = "x86" then
		'wscript.echo "Windows 32bit system detected"
		isSixtyFour=0
	elseif OsType = "AMD64" then
		'wscript.echo "Windows 64bit system detected"
		isSixtyFour=1
	end if
End Function
