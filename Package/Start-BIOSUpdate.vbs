'Run Powershell script of the same basename and hides console (No blink)
'IOW: the name of this VBS script and the name of the Powershell script must match!!

'Get name of this script and replace extension with .PS1
strN0 = Replace(Ucase(WScript.ScriptFullName), ".VBS", ".PS1")

'Read Arguments
Set objArgs = Wscript.Arguments
For Each strArg in objArgs
	If Instr(1, strArg, " ", vbTextCompare)>0 then 'if contains a space
		StrAllArg=StrAllArg&" """&strArg&""""	'Enclose with double quotes
	Else
		StrAllArg=StrAllArg&" "&strArg
	End If
Next

'Define commanline
StrCommand = ("powershell.exe -ExecutionPolicy Bypass -nologo -File """&strN0&"""" & StrAllArg)

'Create object
set WSshell = CreateObject("WScript.Shell")

'Change directory in case the scipt path gets removed (Makes Start-Job fail in powershell)
WSshell.CurrentDirectory = WSshell.ExpandEnvironmentStrings("%TEMP%")

'Launch Powershell hidden
ReturnValue = WSshell.Run(StrCommand,0,true)  '0=run Hidden, true=Wait for exit
WScript.Quit(ReturnValue)