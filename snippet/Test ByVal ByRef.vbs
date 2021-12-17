' Proof of VBS' byRef parameter passing trap/bug/"feature" by Denis St-Pierre
'
' If you like to be consistent and re-use the same variable names in functions
' throughout the same script, you should be aware of this little wrinkle in VBS
'
' Usually when you call a function in MOST languages, you give it the "contents" of a variable
' and if you want you can give it a variable "by Reference" to save memory or gain speed if it's
' contents are large (e.g.: a text file, an array, etc.)
'
' In VBS, however, when you call a function and pass it a variable, it is "by Reference" by default!
'	IOW: You give the function the actual variable to play with (egg. the container) not the Contents.
' This means that by default, all variables are effectively "Public" in scope throughout the script
' EVEN IF YOU DECLARE THE VARIABLE (via Dim) IN THE FUNCTION or SUB
'
' For proof, play around with the "Function DoStuff" line and you will get different results.
'
' WORKAROUNDS:
'	1-Use unique variable names in ALL functions and Subs
'	2-Declare Functions and sub that have inputs variables using ByVal as described below
'
' FYI: ByVal means By Value


Test()
WScript.Quit(0)	



Function Test()
	Dim sString	'By doing this here the variable is now "Private"
			'(or so we are lead to believe)

	sString = "hello "
	MsgBox "sString is =" & sString
	
	Call DoStuff( sString )
	
	MsgBox "sString is =" & sString	'Should still be "hello " if the variable is Private
End Function
	
	
Function DoStuff(sString)		'This is the same as "Function DoStuff(ByRef sString)"
'Function DoStuff(ByVal sString)	'This will make it that sString variable Private
'Function DoStuff(ByRef sString)	'This is the same as "Function DoStuff( sString)"
	sString = sString & sString & sString
	DoStuff = True
End Function