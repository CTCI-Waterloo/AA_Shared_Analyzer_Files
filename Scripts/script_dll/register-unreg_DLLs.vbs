
Set fso = CreateObject("Scripting.FileSystemObject")
Set WSHShell = CreateObject("WScript.Shell")

strCurDir = WSHShell.currentdirectory
strFixDir = chr(34) & "%SYSTEMROOT%\system32\regsvr32.exe" & chr(34) &  " /U" & " /S"

set folderTMP = fso.GetFolder(strCurDir)
Set filesTMP = folderTMP.Files

If filesTMP.Count <> 0 Then
	For Each strfile In filesTMP
		if Ucase(Right(strfile,4)) = ".DLL" then
			shellCall = strFixDir & " " & chr(34) & strfile & chr(34)
			'msgbox shellCall
			Call WSHShell.run(shellCall,3,True)
		end if
	Next
End If
