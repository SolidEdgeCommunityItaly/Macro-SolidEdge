Option Explicit


Class dummy_lib

	Private sub class_Initialize
		' Called automatically when class is created
	End Sub

	Private sub class_Terminate
		' Called automatically when all references to class instance are removed
	End Sub

	Public Sub Test()
		MsgBox "Test is ok", vbOKOnly, "dummy_lib"
	End sub

End Class
