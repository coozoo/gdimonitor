'Script to kill all VBscripts.

Dim strComputer
Dim flag

Set wshShell = CreateObject("WScript.Shell")
strComputer = "."


		Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
			Set colProcessList = objWMIService.ExecQuery _
			("Select * from Win32_Process Where Name = 'cscript.exe' OR Name = 'wscript.exe'")
				For Each objProcess in colProcessList
					objProcess.Terminate()
				Next







