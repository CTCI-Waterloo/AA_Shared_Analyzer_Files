Set fso = CreateObject("Scripting.FileSystemObject")
Set WSHShell = CreateObject("WScript.Shell")

WinNetVersion = "v2.0.50727"
strArguments = " /codebase /verbose "
strDOTNETFramework = "C:\WINDOWS\Microsoft.NET\Framework\"
strREGASM = "\regasm.exe"
strCommand = strDOTNETFramework & WinNetVersion & strREGASM & strArguments

If Not fso.FileExists(strDOTNETFramework & WinNetVersion & strREGASM) Then
	MsgBox "DOT .NET v3.5 is not installed & vbNewLine & Please consult with your IT Technician to have it installed"
	Wscript.Quit
End If 

strCurDir = WSHShell.currentdirectory

arFiles = Array("Analyzer_MPR_Translator_DLL.dll")

For Each strfile In arFiles
	if Ucase(Right(strfile,4)) = ".DLL" then
		shellCall = strCommand & chr(34) & strCurDir & "\" & strfile & chr(34)
		'msgbox shellCall
		Call WSHShell.run(shellCall,3,True)
	end if
Next