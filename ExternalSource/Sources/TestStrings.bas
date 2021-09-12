Attribute VB_Name = "TestStrings"

	' add your procedures here


'@TestMethod("Primitive")
Private Sub Test01_Dedup()
	On Error GoTo TestFail
	
	'Arrange:
	Dim myExpected  As String
	myExpected = "Hello Worlde"
	
	
	Dim myResult As String
	
	'Act:
	myResult = Strings.Dedup("Heeello Worldee", "e")

	'Assert:
	Assert.Exact.AreEqual myExpected, myResult, CurrentProcedureName
	
TestExit:
	Exit Sub
	
TestFail:
	Debug.Print CurrentComponentName & "." & CurrentProcedureName & "  raised an error: #" & Err.Number & " - " & Err.Description
	Resume TestExit
	
End Sub

Private Sub Test02_TrimmerDefault()
	On Error GoTo TestFail
	
	'Arrange:
	Dim myExpected  As String
	myExpected = "Hello World"
	
	
	Dim myResult As String
	
	'Act:
	myResult = Strings.Trimmer("   ;;;,;,;Hello World ;,; ;; ,")

	'Assert:
	Assert.Exact.AreEqual myExpected, myResult, CurrentProcedureName
	
TestExit:
	Exit Sub
	
TestFail:
	Debug.Print CurrentComponentName & "." & CurrentProcedureName & "  raised an error: #" & Err.Number & " - " & Err.Description
	Resume TestExit
	
End Sub