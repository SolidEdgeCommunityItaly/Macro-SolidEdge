' Per gestire meglio gli errori è bene interrompere il codice con "Exit Sub" o "Exit Function"
' Quindi trovo più semplice rinchiudere lo script in una Sub (es.: Sub Main() oppure Sub NomeDelloScript())
' quindi a inizio file faccio partire il tutto con un "Call Main()" oppure "Call NomeDelloScript()"

Call Main()

Sub Main()
	Dim Message, Title
	Title = Left(Wscript.ScriptName, Len(Wscript.ScriptName) - 4) ' script name
	
	' il codice che segue potrebbe generare errori ("eccezioni")
	' quindi lo inserisco in un costrutto <On Error Resume Next>...<On Error GoTo 0>

	On Error Resume Next
		'nb: in vbscript le label non funzionano
		
		Call funzione_che_genera_errori() '(questa ad esempio non esiste)
		stop
		If Err.Number = 5 Then
			' Errore noto
			MsgBox "Attenzione: qui scrivere che cosa deve fare l'utente per evitare questo errore o cosa conosco di questo errore.", vbInformation, Title
			Err.Clear
			' Azioni/funzioni da compiere per correggere o Exit Sub / Exit Function
			'...
			
		ElseIf Err.Number = 6 Then
			' eccetera
			Err.Clear
			
		ElseIf Err Then
			MsgBox "Nuovo errore/eccezione da imparare a gestire!" & vbCrLf & vbCrLf &_
				"Number: " & Err.Number & vbCrLf &_
				"Description: " & Err.Description & vbCrLf &_
				"Source: " & Err.Source & vbCrLf &_
				"HelpContext: " & Err.HelpContext & vbCrLf &_
				"HelpFile: " & Err.HelpFile & vbCrLf &_
				"" & vbCrLf & vbCrLf & "Esco con Exit Sub." &_
				"", vbInformation, Title
			Exit Sub
		End If

	' ripristina la gestione degli errori, così si troveranno nuovi bug 
	' e si eviteranno comportamenti indesiderati
	On Error GoTo 0

	MsgBox("Fatto.")

End Sub
