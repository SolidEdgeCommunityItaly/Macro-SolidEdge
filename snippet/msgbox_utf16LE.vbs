MsgBox "Test codifica del file in ""UTF-16 LE"" (""UCS-2 LE BOM"") anzichè UTF-8: àèìòù", vbOKOnly + vbQuestion, "Test MsgBox in UTF-16 LE"

' Test with function ExecuteGlobal
ExecuteGlobal "MsgBox(""àèìòù"")"