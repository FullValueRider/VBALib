Attribute VB_Name = "TestTrackedXY"
''@IgnoreModule
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module

'Private Assert As Object
'Private Fakes As Object

#If twinbasic Then
    'Do nothing
#Else


    '@ModuleInitialize
Private Sub ModuleInitialize()
    GlobalAssert
    'this method runs once per module.
    ' Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub


''@TestInitialize
'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'End Sub
'
'
''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub


#End If

Public Sub TrackedXYTests()
 
    #If twinbasic Then
        Debug.Print CurrentProcedureName;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
    #End If

    Test01_SeqObj
    Test02_Initialised
    Test03_Move_South_3
    
    Debug.Print vbTab, vbTab, vbTab, "Testing completed"

End Sub


'@TestMethod("SeqA")
Private Sub Test01_SeqObj()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "TrackedXY", "TrackedXY")
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myResult(0) = VBA.IsObject(myT)
    myResult(1) = VBA.TypeName(myT)
    myResult(2) = myT.TypeName
    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqA")
Private Sub Test02_Initialised()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
    Dim myExpected As Variant
    myExpected = Array("0,0", False, True, e_Heading.m_North)
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myT.Location.ToString
    myResult(1) = myT.BoundsInUse
    myResult(2) = myT.AtOrigin
    myResult(3) = myT.Heading
    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqA")
Private Sub Test03_Move_South_3()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
    Dim myExpected As Variant
    myExpected = Array("0,-3", True, False, e_Heading.m_South, "{{0,0},{0,-1},{0,-2},{0,-3}}")
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myT.Move "South", 3
'    myresult(0) = myT.Location.ToString
'    myresult(1) = myT.Moved
'    myresult(2) = myT.AtOrigin
'    myresult(3) = myT.Heading
'    myresult(4) = Fmt.Text("0", myT.Trail)
    
    Debug.Print myT.Location.ToString
    Debug.Print myT.Moved
    Debug.Print myT.AtOrigin
    Debug.Print myT.Heading
    Debug.Print Fmt.Text("0", myT.Trail)
    
    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
