MsgBox "Test codifica del file in UTF-8 anzichè ANSI: àèìòù", vbOKOnly + vbQuestion, "Test MsgBox in UTF-8"


' Test with function ExecuteGlobal
ExecuteGlobal "MsgBox(""àèìòù"")"