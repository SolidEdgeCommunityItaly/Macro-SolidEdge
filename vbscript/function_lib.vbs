Option Explicit

' Common Fuctions Library - vespadj

' How include this function lib:
' see and copy from "function_lib_execute_test.vbs"

If Wscript.ScriptName = "function_lib.vbs" Then
	Dim answer
	answer = MsgBox ("Questo script è una libreria, non lanciare direttamente." & vbCrLf & "'Yes' per chiudere, 'No' per eseguire i test di funzionamento." &_
	vbCrLf & vbCrLf &_
	"This script is a Library: don't run it directly." & vbCrLf & "'Yes' to close, 'No' for launch internal testing." &_
	"" , vbYesNo, "function_lib.vbs")
	
	If answer = 7 Then
		' Test
		Dim this_lib
		Set this_lib = New function_lib
		
		MsgBox "Test GetExtensionFromFilename(""ABC123456.pdf""): " & vbcrlf & this_lib.GetExtensionFromFilename("ABC123456.pdf") , vbOKOnly, "function_lib.vbs"
		MsgBox "Test GetExtensionFromFilename(""noextension""): " & vbcrlf & this_lib.GetExtensionFromFilename("noextension") , vbOKOnly, "function_lib.vbs"
		MsgBox "Test RegExpTest ok: "   & vbcrlf & this_lib.RegExpTest("^[0-9A-Z]{9}(_[^_]+)?(\.(pdf|zip))", "ABC123456.pdf") , vbOKOnly, "function_lib.vbs"
		MsgBox "Test RegExpTest null: " & vbcrlf & this_lib.RegExpTest("^[0-9A-Z]{9}(_[^_]+)?(\.(pdf|zip))", "zABC123456.pdf") , vbOKOnly, "function_lib.vbs"
		
		Set this_lib = Nothing
	End If
End If

Class function_lib
	Private Sub class_Initialize
		' Called automatically when class is created
		
		' === CONSTANTS - Costanti ===
		' Constant not allowed in Class
		'https://stackoverflow.com/questions/21052084/constant-inside-class
		
		' SolidEdgeConstants.ModelMemberComponentTypeConstants
		Dim seShowPart, seHidePart, seSectionPart, seUndefinedDisplay 
		Dim seSolidBodyMemberType, seTubeCenterlineMemberType
		Dim seOccurrences
		seShowPart = 0
		seHidePart = 1
		seSectionPart = 2
		seUndefinedDisplay = 3
		seSolidBodyMemberType  = 7
		seTubeCenterlineMemberType = 10
		
		'SolidEdgeFramework.ObjectType
		seOccurrences = -825730197
		
	End Sub
	
	
	Private Sub class_Terminate
		' Called automatically when all references to class instance are removed
	End Sub
	
	
	Public Sub Test()
		'Wscript.Echo("Test is OK!")
		MsgBox "Test is OK!", vbOKOnly, "Include function_lib.vbs"
	End Sub
	
	
	' === STRINGS - Stringhe di Testo ===
	
	'non completamente verificato
	Public Function RegExpTest(ByVal pattern, ByVal strng)
		Dim regEx, Match, Matches   ' Create variable.
		Set regEx = New RegExp   ' Create a regular expression.
		regEx.Pattern = pattern   ' Set pattern.
		regEx.IgnoreCase = True   ' Set case insensitivity.
		regEx.Global = True   ' Set global applicability.
		Set Matches = regEx.Execute(strng)   ' Execute search.
		' DEBUG:
		' Msgbox (Matches.Count)
		' Dim RetStr
		' For Each Match in Matches   ' Iterate Matches collection.
		' 	RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
		' 	RetStr = RetStr & Match.Value  ' non so perchè ma ne trova 2 e il secondo è vuoto
		' 	'RetStr = Match.Value 
		' Next
		' Msgbox(RetStr)
		If Matches.Count > 0 Then
			RegExpTest = Matches.Item(0).Value ' 0 = First
		End If
	End Function
	
	
	' === ARRAYS ===
	
	Public Function SortArray(ByVal MyArray , ByVal order)
		'was
		'Function SortArray(MyArray() As Variant)
		If order ="" Then
			order = "asc"
		End If
		Dim First  'As Integer
		Dim Last   'As Integer
		Dim i      'As Integer
		Dim j      'As Integer
		Dim Temp   'As String
		Dim List   'As String
		
		SortArray = ""
		First = LBound(MyArray)
		Last = UBound(MyArray)
		order = LCase(order)
		For i = First To Last - 1
			For j = i + 1 To Last
				If order = "asc" Then 'ascendente
					If MyArray(i) > MyArray(j) Then
						Temp = MyArray(j)
						MyArray(j) = MyArray(i)
						MyArray(i) = Temp
					End If
				ElseIf order = "desc" Then 'descrescente then
					If MyArray(i) < MyArray(j) Then
						Temp = MyArray(j)
						MyArray(j) = MyArray(i)
						MyArray(i) = Temp
					End If
				End If
			Next
		Next
		
		'For i = 1 To UBound(MyArray)
		'    List = List & vbCrLf & MyArray(i)
		'Next
		'MsgBox List
		SortArray = MyArray
		
	End Function
	
	
	' === OBJECTS ===
	'http://stackoverflow.com/questions/250970/detect-a-error-object-doesnt-support-this-property-or-method
	Public Function SupportsMember(ByVal object, ByVal memberName)
		On Error Resume Next
		
		Dim x
		Eval("x = object." + memberName)
		
		If Err = 438 Then
			SupportsMember = False
		Else
			SupportsMember = True
		End If
		
		On Error GoTo 0 'clears error
	End Function
	
	
	' === FILES ===
	Public Function file_exist(ByVal filename)
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		file_exist = fso.FileExists(filename)
		Set fso = Nothing
	End Function
	
	'Function GetFilenameFromPath(ByVal strPath As String) As String
	' Returns the rightmost characters of a string upto but not including the rightmost '\'
	' e.g. 'c:\winnt\win.ini' returns 'win.ini'
	Public Function GetFilenameFromPath(ByVal strPath)
		If Right(strPath, 1) <> "\" And Len(strPath) > 0 Then
			GetFilenameFromPath = GetFilenameFromPath(Left(strPath, Len(strPath) - 1)) + Right(strPath, 1)
		End If
	End Function
	
	Public Function GetExtensionFromFilename(ByVal strPath)
		' If Right(strPath, 1) <> "." And Len(strPath) > 0 Then
		' 	GetExtensionFromFilename = GetExtensionFromFilename(Left(strPath, Len(strPath) - 1)) + Right(strPath, 1)
		' End If
		Dim filename, ext
		filename = GetFilenameFromPath(strPath)
		ext = Right(filename, Len(filename) - InStrRev(filename, "."))
		If ext = filename Then
			GetExtensionFromFilename = ""
		Else
			GetExtensionFromFilename = ext
		End If
	End Function
	
	
	Private Function ReadTextFile(ByVal filepath)
		Dim fso, f, s
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set f = fso.OpenTextFile(filepath, 1) ' in Read Only
		s = f.ReadAll()
		f.Close 
		ReadTextFile = s
	End Function
	
	
	'http://www.robvanderwoude.com/vbstech_files_ini.php
	Public Function ReadIni(ByVal myFilePath, ByVal mySection, ByVal myKey)
		' This function returns a value read from an INI file
		'
		' Arguments:
		' myFilePath  [string]  the (path and) file name of the INI file
		' mySection   [string]  the section in the INI file to be searched
		' myKey       [string]  the key whose value is to be returned
		'
		' Returns:
		' the [string] value for the specified key in the specified section
		'
		' CAVEAT:     Will return a space if key exists but value is blank
		'
		' Written by Keith Lacelle
		' Modified by Denis St-Pierre and Rob van der Woude
		
		Const ForReading   = 1
		Const ForWriting   = 2
		Const ForAppending = 8
		
		Dim intEqualPos
		Dim objFSO, objIniFile
		Dim strFilePath, strKey, strLeftString, strLine, strSection
		
		Set objFSO = CreateObject( "Scripting.FileSystemObject" )
		
		ReadIni     = ""
		strFilePath = Trim( myFilePath )
		strSection  = Trim( mySection )
		strKey      = Trim( myKey )
		
		If objFSO.FileExists( strFilePath ) Then
			Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
			Do While objIniFile.AtEndOfStream = False
				strLine = Trim( objIniFile.ReadLine )
				
				' Check if section is found in the current line
				If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
					strLine = Trim( objIniFile.ReadLine )
					
					' Parse lines until the next section is reached
					Do While Left( strLine, 1 ) <> "["
						' Find position of equal sign in the line
						intEqualPos = InStr( 1, strLine, "=", 1 )
						If intEqualPos > 0 Then
							strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
							' Check if item is found in the current line
							If LCase( strLeftString ) = LCase( strKey ) Then
								ReadIni = Trim( Mid( strLine, intEqualPos + 1 ) )
								' In case the item exists but value is blank
								If ReadIni = "" Then
									ReadIni = " "
								End If
								' Abort loop when item is found
								Exit Do
							End If
						End If
						
						' Abort if the end of the INI file is reached
						If objIniFile.AtEndOfStream Then Exit Do
						
						' Continue with next line
						strLine = Trim( objIniFile.ReadLine )
					Loop
					Exit Do
				End If
			Loop
			objIniFile.Close
		Else
			WScript.Echo strFilePath & " doesn't exists. Exiting..."
			Wscript.Quit 1
		End If
	End Function
	
	Public Sub WriteIni(ByVal myFilePath, ByVal mySection, ByVal myKey, ByVal myValue)
		' This subroutine writes a value to an INI file
		'
		' Arguments:
		' myFilePath  [string]  the (path and) file name of the INI file
		' mySection   [string]  the section in the INI file to be searched
		' myKey       [string]  the key whose value is to be written
		' myValue     [string]  the value to be written (myKey will be
		'                       deleted if myValue is <DELETE_THIS_VALUE>)
		'
		' Returns:
		' N/A
		'
		' CAVEAT:     WriteIni function needs ReadIni function to run
		'
		' Written by Keith Lacelle
		' Modified by Denis St-Pierre, Johan Pol and Rob van der Woude
		
		Const ForReading   = 1
		Const ForWriting   = 2
		Const ForAppending = 8
		
		Dim blnInSection, blnKeyExists, blnSectionExists, blnWritten
		Dim intEqualPos
		Dim objFSO, objNewIni, objOrgIni, wshShell
		Dim strFilePath, strFolderPath, strKey, strLeftString
		Dim strLine, strSection, strTempDir, strTempFile, strValue
		
		strFilePath = Trim( myFilePath )
		strSection  = Trim( mySection )
		strKey      = Trim( myKey )
		strValue    = Trim( myValue )
		
		Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
		Set wshShell = CreateObject( "WScript.Shell" )
		
		strTempDir  = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
		strTempFile = objFSO.BuildPath( strTempDir, objFSO.GetTempName )
		
		Set objOrgIni = objFSO.OpenTextFile( strFilePath, ForReading, True )
		Set objNewIni = objFSO.CreateTextFile( strTempFile, False, False )
		
		blnInSection     = False
		blnSectionExists = False
		' Check if the specified key already exists
		blnKeyExists     = ( ReadIni( strFilePath, strSection, strKey ) <> "" )
		blnWritten       = False
		
		' Check if path to INI file exists, quit if not
		strFolderPath = Mid( strFilePath, 1, InStrRev( strFilePath, "\" ) )
		If Not objFSO.FolderExists ( strFolderPath ) Then
			WScript.Echo "Error: WriteIni failed, folder path (" _
			& strFolderPath & ") to ini file " _
			& strFilePath & " not found!"
			Set objOrgIni = Nothing
			Set objNewIni = Nothing
			Set objFSO    = Nothing
			WScript.Quit 1
		End If
		
		While objOrgIni.AtEndOfStream = False
			strLine = Trim( objOrgIni.ReadLine )
			If blnWritten = False Then
				If LCase( strLine ) = "[" & LCase( strSection ) & "]" Then
					blnSectionExists = True
					blnInSection = True
				ElseIf InStr( strLine, "[" ) = 1 Then
					blnInSection = False
				End If
			End If
			
			If blnInSection Then
				If blnKeyExists Then
					intEqualPos = InStr( 1, strLine, "=", vbTextCompare )
					If intEqualPos > 0 Then
						strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )
						If LCase( strLeftString ) = LCase( strKey ) Then
							' Only write the key if the value isn't empty
							' Modification by Johan Pol
							If strValue <> "<DELETE_THIS_VALUE>" Then
								objNewIni.WriteLine strKey & "=" & strValue
							End If
							blnWritten   = True
							blnInSection = False
						End If
					End If
					If Not blnWritten Then
						objNewIni.WriteLine strLine
					End If
				Else
					objNewIni.WriteLine strLine
					' Only write the key if the value isn't empty
					' Modification by Johan Pol
					If strValue <> "<DELETE_THIS_VALUE>" Then
						objNewIni.WriteLine strKey & "=" & strValue
					End If
					blnWritten   = True
					blnInSection = False
				End If
			Else
				objNewIni.WriteLine strLine
			End If
		Wend
		
		If blnSectionExists = False Then ' section doesn't exist
			objNewIni.WriteLine
			objNewIni.WriteLine "[" & strSection & "]"
			' Only write the key if the value isn't empty
			' Modification by Johan Pol
			If strValue <> "<DELETE_THIS_VALUE>" Then
				objNewIni.WriteLine strKey & "=" & strValue
			End If
		End If
		
		objOrgIni.Close
		objNewIni.Close
		
		' Delete old INI file
		objFSO.DeleteFile strFilePath, True
		' Rename new INI file
		objFSO.MoveFile strTempFile, strFilePath
		
		Set objOrgIni = Nothing
		Set objNewIni = Nothing
		Set objFSO    = Nothing
		Set wshShell  = Nothing
	End Sub
	
	
	' === SOLID EDGE===
	Public Function OpenSolidEdgeDoc(ByVal FullFilePath)
		'return objDoc 
		'usage:
		' Dim objDoc
		' Set objDoc = OpenSolidEdgeDoc(FullFilePath)
		' 'code...
		' objDoc.Close([SaveChanges])
		' Set objDoc = Nothing
		
		Dim objApp ' As SolidEdgeFramework.Application
		Dim objDoc
		
		' Create/get the application with specific settings
		On Error Resume Next
		Set objApp = GetObject(, "SolidEdge.Application")
		If Err Then
			Err.Clear
			Set objApp = CreateObject("SolidEdge.Application")
		End If
		On Error Resume Next
		
		Set objDoc = objApp.Documents.Open(FileMaster)
		objApp.Visible = True
		
		OpenSolidEdgeDoc = objDoc
		
		Set objApp = Nothing
		Set objDoc = Nothing
	End Function
	
	
	' === FILES PROPERTIES (SOLID EDGE)===
	' === NOTA BENE: Da SolidEdge ST10, 
	' === dopo l'installazione principale e di ogni Maintenance Pack (MP)
	' === SolidEdgeFileProperties "PropAuto.dll" va ri-registrato
	' === con una utility che non è presente in questo reposity (non ho verificato il copyright)
	Public Function FilePropGet(ByVal FullFilePath, ByVal Scheda, ByVal Name)
		If Scheda = "" Then
			Scheda = "Custom"
		End If
		
		Dim objPropertySets 'As SolidEdgeFileProperties.PropertySets = Nothing
		Dim objProperties	 'As SolidEdgeFileProperties.Properties = Nothing
		
		' Create a new instance of PropertySets.
		' objPropertySets = New SolidEdgeFileProperties.PropertySets
		Set objPropertySets = CreateObject("SolidEdge.FileProperties")
		
		Call objPropertySets.Open(FullFilePath, True) ' readonly
		Set objProperties = objPropertySets.Item(Scheda)
		
		If Name<>"" Then
			FilePropGet = objProperties(Name).Value
			'msgbox(FilePropGet)
			'TODO: if Name = "" return an array{Name="...", Value="..."} with all Proprieties
		End If
		
		objPropertySets.Close()
		
		Set objProperties = Nothing
		Set objPropertySets = Nothing
		
	End Function
	
	
	Public Sub FilePropSet(ByVal File, ByVal Scheda, ByVal name, ByVal value, ByVal ProcessFamilyOfAssembly)
		'Todo: fare un'altra funzione dove al posto di chiamare la funzione per ogni campo-valore, passo direttamente una matrice di campi e valori
		' dovrei così impiegare meno tempo infatti ora esegue un -apri file, salva, chiudi- per ogni prop sullo stesso file
		Dim objProps 	'As Object
		Dim objProp 	'As Object
		Dim FamAsm 		'As Boolean
		Dim MemberCount 'As Long
		Dim MemberNames '() 'As String
		
		'dbg
		'File = "\\Server\NR\A\000001\NRA000001.asm"
		'Name = "Codice"
		'value = "NRA000001"
		'Scheda = "Custom"
		'ProcessFamilyOfAssembly = True
		
		' NB: Prima di usare queste funzioni assicurarsi se stai lavorando con un assieme con Famiglia di Assiemi
		
		Set objProps = CreateObject("SolidEdge.FileProperties")
		
		'FamAsm = False
		''msgbox(TypeName(FamAsm))
		'Call objProps.IsFileFamilyOfAssembly(File, FamAsm)
		'If FamAsm Then
		'    Call objProps.GetFamilyOfAssemblyMemberNames(File, MemberCount, MemberNames)
		'    If Not ProcessFamilyOfAssembly Then
		'        MsgBox ("Attenzione: l'assieme contiente una Famiglia di assimi. N° membri: " & MemberCount & vbCrLf & "Verrano saltati, per processarli chiamare la funzione con" & vbCrLf & '"ProcessFamilyOfAssembly = True ")
		'    Else
		'    
		'    End If
		'    For i = 1 To MemberCount Step 1
		'        Call objProps.Open(File & "!" & MemberNames(i))
		'        Set objProp = objProps(Scheda)
		'        Call objProp.Add(name, value)
		'        'Call objProps.Save
		'        Call objProps.Close
		'    Next
		'End If
		
		Call objProps.Open(File)
		Set objProp = objProps(Scheda)
		
		On Error Resume Next 'GoTo metodo2 !!! in vbscript le label non funzionano
		' metodo1: vale per scheda Custom
		Call objProp.Add(name, value)
		On Error GoTo 0
		'metodo2:
		' metodo2: vale per schede SummaryInformation, Projectinformation,...
		objProp(name).value = value
		
		Call objProps.Save
		Call objProps.Close
		
		Set objProps = Nothing
		Set objProp = Nothing
		
	End Sub
	
	
	Public Function SetHardwareFile(ByVal FullFileName, ByVal value)
		Dim objProps    'As SolidEdgeFileProperties.Properties
		Dim objProp     'As Solidedgefileproperties.Property
		Title = "SetHardwareFile"
		'  Se il file è in sola lettura, toglie readonly
		' verrà di seguito reimpostato
		' TODO: creare l'oggeto FileSystem
		'If FileSystem.GetAttr(FullFileName) = (33 Or 1) Then
		'	Call MsgBox("file in sola lettura; interrompere a mano VBA x uscire, OK per procedere", Title)
		'	fileReadOnly = True
		'	Call FileSystem.SetAttr(FullFileName, 32)
		'Else
		'	fileReadOnly = False
		'End If
		
		' core
		Set objProps = CreateObject("SolidEdge.FileProperties")
		Call objProps.Open(FullFileName)
		Set objProp = objProps("ExtendedSummaryInformation")
		'On Error Resume Next
		objProp("Hardware").value = value
		'nb: Call objProp.Add(Name, Value) non va
		Call objProps.Save
		' Release objects
		Call objProps.Close
		Set objProps = Nothing
		Set objProp = Nothing
		
		' TODO: FileSystem
		'If fileReadOnly = True Then
		'    Call FileSystem.SetAttr(FullFileName, 33)
		'End If
		
	End Function
	
	' === MATH ===
	
	
	' === OPERATORS ===
	Public Function OrElse(ByVal a, ByVal b)
		Select Case True
			Case a, b
			OrElse = True
			Case Else
			OrElse = False
		End Select
	End Function
	'Test
	'Wscript.echo("Test: OrElse(true, Nothing)" & vbCrLf & "returns: " & OrElse(true, Nothing))
	
	
	
	' === SIGNALS ===
	
	' === SQL ===
	Public Function SqlToArray(ByVal strConn, ByVal fileSql, ByVal params)
		' Return Array( 0: Fields ; 1: Rows(col, row) )
		
		SqlToArray = Array( Array(""), Array("") )	 'init
		
		Dim sql, param
		Dim i
		
		sql = ReadTextFile(fileSql)
		' e.g. of sql file:  select * from table where q1 = '{0}' and q2 = '{1}' ... ;
		
		i=0
		For Each param In params 
			param = Replace(param,"'","''") ' escape apos
			param = Replace(param, "*", "%")
			param = Replace(param, "?", "_")
			If i=0 Then
				sql = Replace(sql, "<TAG>", param)
				sql = Replace(sql, "{0}", param)
			ElseIf i=1 Then
				sql = Replace(sql, "<TAG2>", param)
				sql = Replace(sql, "{1}", param)
			Else
				sql = Replace(sql, "{"&i&"}", param)
			End If
			
			i = i + 1
		Next
		
		'msgbox(strConn)
		'msgbox(sql)
		
		Dim dbConn
		Set dbConn = CreateObject("ADODB.Connection")
		dbConn.Open strConn
		Dim rs
		'Set rs = dbConn.Execute(sql)
		Set rs = CreateObject("ADODB.Recordset")
		Set rs.ActiveConnection = dbConn
		rs.CursorType = 3 'adOpenStatic
		rs.Open sql
		'found = rs.RecordCount
		
		'msgbox("dbConn.State:" & dbConn.State) ' return 0: close; 1: Open
		'msgbox("rs.State:" & rs.State)			' return 0: close; 1: Open
		
		If rs.State = 1 Then				' questo dovrebbe consentire insert e update senza restituzione di recordset
			If Not (rs.EOF Or rs.BOF) Then	' se non è stata restituita una tabella vuota
				Dim arrayFields()
				ReDim arrayFields( rs.Fields.Count-1 ) ' nuovo indice massimo
				For i = 0 To rs.Fields.Count-1
					arrayFields(i) = rs(i).Name
				Next
				SqlToArray = Array( arrayFields, rs.GetRows() )
			End If
		End If
		
		If dbConn.State = 1 Then
			dbConn.Close
		End If
		
		Set dbConn = Nothing
		Set rs = Nothing
		
	End Function
	
	'Funzione di debug che mostra una msgbox contenente i dati per capire quali dati ho. usare per recordset piccoli
	Public Sub SqlToArray_dbg(ByVal arrayTab)
		Dim campi, dati
		campi = arrayTab(0)
		dati  = arrayTab(1)
		If IsArray(dati) Then
			txt = Join(campi, " | ") & vbcrlf
			For r = LBound(dati, 2) To UBound(dati, 2)
				For c = LBound(dati, 1) To UBound(dati, 1)
					txt = txt & dati(c, r) & " | "
				Next
				txt = Left(txt, Len(txt) - 3) & vbcrlf
			Next
			MsgBox txt, vbOKOnly, "SqlToArray_dbg"
		End If
	End Sub
	
	' === TIME ===
	
	
	
End Class

