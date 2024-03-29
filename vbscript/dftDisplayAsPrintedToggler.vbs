' (en) Macro for Draft: Toggle Display as Printed
' (it) Macro per Draft: alterna "Visualizza come stampato"

' v.1.1	2020/04/13 First Realise

Call dftDisplayAsPrintedToggler()

Sub dftDisplayAsPrintedToggler()
	Dim Message, Title
		Title = Left(Wscript.ScriptName, Len(Wscript.ScriptName) - 4) ' script name

	Dim objApp 'As SolidEdgeFramework.Application
	Dim objDoc 'As SolidEdgeDraft.DraftDocument

	' Create/get the application with specific settings
	On Error Resume Next
		Set objApp = GetObject(, "SolidEdge.Application")
		objApp.Visible = True
	If Err Then
		Err.Clear
		MsgBox "Prima apri un Draft in Solid Edge." & vbCrLf & vbCrLf & "Open a Draft in Solid Edge first.", vbInformation , Title
		Exit Sub
	Else
		Set objDoc = objApp.ActiveDocument
	End If
	On Error GoTo 0   'rigestisce gli errori

	If objApp.ActiveEnvironment = "Detail" Or objApp.ActiveEnvironment = "DrawingViewEdit" Or objApp.ActiveEnvironment = "TwoDModel" Then
		objApp.ActiveWindow.DisplayAsPrinted = Not objApp.ActiveWindow.DisplayAsPrinted 
		objApp.StartCommand (32876) ' DetailViewRefreshWindows
	Else
		MsgBox "Prima apri un Draft in Solid Edge." & vbCrLf & vbCrLf & "Open a Draft in Solid Edge first.", vbInformation , Title
	End If

	' Release objects
	Set objApp = Nothing
	Set objDoc = Nothing
End Sub
