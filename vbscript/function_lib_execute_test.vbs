Option Explicit

' Importing External Script Code
' To Run in VS Code, prefix library file with curPath: the path of this script

Private Function Include(ByVal vbsFile)
	Dim fso, f, s
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(vbsFile, 1) ' in Read Only
	s = f.ReadAll()
	f.Close 
	ExecuteGlobal s
End Function

Dim curPath ' As String
curPath = Left(Wscript.ScriptFullName, InStrRev(Wscript.ScriptFullName, "\") -1 )
' Wscript.Echo(curPath)

Include(curPath & "\" & "function_lib.vbs")

Dim Lib
Set Lib = New function_lib

' test
Lib.Test()
MsgBox(Lib.GetExtensionFromFilename("function_lib.vbs"))