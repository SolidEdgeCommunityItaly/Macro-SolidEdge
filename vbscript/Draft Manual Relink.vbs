Option Explicit

Call Main()

Sub Main()
    ' Global Settings
    Dim Message, Title
    Title = "Draft Manual Relink"

    Dim rmApp, rmDFT, rm3D, link
    Dim q, inputDFT, inputLink

    On Error Resume Next
        Set rmApp = GetObject(, "RevisionManager.Application")
        If rmApp Is Nothing Then
            Set rmApp = CreateObject("RevisionManager.Application")
        End If
    On Error GoTo 0

    Message = "ATTENZIONE! Funzione senza supervisione degli errori, sovrascrittura diretta dei link del Draft." &_
        vbCrLf & "Annullare se non si ha consapevolezza dell'operazione." &_
        vbCrLf & "Caratteri accentati nei file (ANSI/utf-8) non supportati." &_
        vbCrLf & vbCrLf &_
        vbCrLf & "Utilizzare Design Manager, se possibile." &_
        vbCrLf & vbCrLf &_
        "Immettere il percorso completo del file dft:"
    q = InputBox(Message, Title, "")
    
    If q="" Then
        Exit Sub
    End If

    inputDFT = replace(q, """", "")

    Set rmDFT = rmApp.Open (inputDFT )
    'Set rm3D = rmDFT.LinkedDocuments.Item(1)

    If rmDFT.LinkedDocuments.Count = 0 Then
		MsgBox "Questo Draft non contiente link a 3D." &_
		vbCrLf & inputDFT, vbOKOnly, Title
	End If
	
	For Each link in rmDFT.LinkedDocuments
        Message = "Link originale: " &_
			vbCrLf & link &_
			vbCrLf &_
			vbCrLf & "Immettere il percorso completo o relativo del file 3D, speficifare la variante se necessario:"
			
        q = InputBox(Message, Title, "")
        inputLink =  replace(q, """", "")
        
        If inputLink="" Then
            Exit Sub
        End If
        
        'msgbox(q)
        link.Replace (inputLink)
    Next

    rmDFT.SaveAllLinks()

    MsgBox "Finito!" , vbOKOnly, Title

    Set rmDFT = Nothing
    Set rmApp = Nothing
End Sub