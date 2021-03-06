' (en) Macro for Draft: Toggle Auto Fill Background of Dimensions, Balloons, ect...
' (it) Riempi Testo con colore di Sfondo sulla Quota/e o Richiamo/i selezionato
' v.1.1	2020/04/13 First Realise

Call dftTextAutoFillBackground()

Sub dftTextAutoFillBackground()
	Dim objApp 'As SolidEdgeFramework.Application
	Dim objDoc 'As SolidEdgeDraft.DraftDocument
	Dim objSel

	' Create/get the application with specific settings
	On Error Resume Next
		Set objApp = GetObject(, "SolidEdge.Application")
		objApp.Visible = True
	If Err Then
		Err.Clear
		MsgBox "Prima apri Solid Edge e seleziona qualcosa." & vbCrLf & vbCrLf & "Open a Draft first and select something.", vbInformation, Title
		Exit Sub
	Else
		Set objDoc = objApp.ActiveDocument
	End If
	On Error GoTo 0   'rigestisce gli errori

	If objApp.ActiveEnvironment = "Detail" Or objApp.ActiveEnvironment = "DrawingViewEdit" Or objApp.ActiveEnvironment = "TwoDModel" Then
		Set objSel = objApp.ActiveSelectSet
		If objSel.Count > 0 Then
			
			For Each obj In objSel
				
				' If object type is: 
				' (en) Dimension, Callout or Balloon, SurfaceFinishSymbol, FeatureControlFrame, DatumFrame
				' (it) quota, richiamo e pallino, rugosità, toll geom, riferimento -A-,

				' ToDo: TextBox (/casella testo) obj.Type = 2004510816 (needs another methods)
				
				'MsgBox (obj.Type)
				If obj.Type = 488188096 Or obj.Type = 384307874 Or obj.Type = 1546072208 Or obj.Type = 77832960 Or obj.Type = -1727514096  Then

					curValue = obj.Style.TextAutoFillBackground
					obj.Style.TextAutoFillBackground = Not curValue
				End If
			Next
			
		End If
	Else
		MsgBox "Prima apri un Draft in Solid Edge." & vbCrLf & vbCrLf & "Open a Draft in Solid Edge first.", vbInformation , Title
	End If

	' Release objects
	Set objApp = Nothing
	Set objDoc = Nothing
	Set objSel = Nothing

End Sub
