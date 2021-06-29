
Function OrElse(ByVal a, ByVal b)
	Select Case true
	Case a, b
		OrElse = True
	Case Else
		OrElse = False
	End Select
End Function

'Test
Wscript.echo("Test: OrElse(true, Nothing)" & vbCrLf & "returns: " & OrElse(true, Nothing))
'Wscript.echo("GetLocale: " & GetLocale())