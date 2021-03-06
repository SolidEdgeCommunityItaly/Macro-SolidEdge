' Importing External Script Code
' To Run in VS Code, prefix library file with curPath: the path of this script

Private Function Include(ByVal vbsFile)
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


Call ProcessaSelezionati()

Sub ProcessaSelezionati()
	Dim objApp 'As SolidEdgeFramework.Application
	Dim objDoc 'As solidedgeAssembly.AssemblyDocument

	Dim objSel 'As SolidEdgeFramework.SelectSet
	Dim objComp 'As solidedgeAssembly.Occurrence

	Dim AList 'As Variant
	Dim logs
	logs = ""

	Title = "ProcessaSelezionati" ' & " - " & action

	' Create/get the application with specific settings
	'On Error Resume Next
	Set objApp = GetObject(, "SolidEdge.Application")
	objApp.Visible = True
	If Err Then
		Err.Clear
		MsgBox "Prima apri solid edge!!!" , vbOKOnly, Title
		'TO DO: 'controlli
	Else
		Set objDoc = objApp.ActiveDocument

		' Get the active Selection
		Set objSel = objApp.ActiveSelectSet
		If objSel.Count < 1 Then
			MsgBox "Prima di lanciare la macro, Selezionare una o più parti." , vbOKOnly, Title
		Else
			For Each objComp In objSel
				Dim c, d, FullFilePath
				'msgbox objComp.Object.OccurrenceFileName
				FullFilePath = objComp.Object.OccurrenceFileName
				c = Lib.FilePropGet(FullFilePath, "Custom", "Codice")
				d = Lib.FilePropGet(FullFilePath, "Custom", "Descrizione Completa")

				logs = logs & vbCrLf & c & vbTab & d
			Next

			' Debug
			MsgBox "=== Fine ===" & vbCrLf & "Result:" & vbCrLf & logs , vbOKOnly, Title
			logs = ""

		End If

	End If
End Sub
