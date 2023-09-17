Attribute VB_Name = "TestcHashC"
'@IgnoreModule
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module

'Private Assert As Object
'Private Fakes As Object

#If TWINBASIC Then
    'Do nothing
#Else


    '@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module
    GlobalAssert
    '    Set Assert = CreateObject("Rubberduck.AssertClass")
    '    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


#End If

Public Sub cHashCTests()

    
    #If TWINBASIC Then
        Debug.Print CurrentProcedureName; vbTab, vbTab,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    #End If

    
    
    Debug.Print "Testing completed"

End Sub


'@TestMethod("cHashC")
Private Sub Test01_SeqObj()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myH As cHashC
    Set myH = New cHashC
    Dim myExpected As Variant
    myExpected = Array(True, "cHashC", "cHashC")
    
    Dim myresult(0 To 2) As Variant
    
    'Act:
    myresult(0) = VBA.IsObject(myH)
    myresult(1) = VBA.TypeName(myH)
    myresult(2) = myH.TypeName
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("cHashC")
Private Sub Test02a_Add_MultipleItems_Count()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As cHashC
    Dim myExpected As Variant
    myExpected = 3&

    Dim myresult As Variant

    'Act:
    Set myH = New cHashC
    myH.Add 42
    myH.Add "Hello"
    myH.Add 3.142
    
    myresult = myH.Count
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("cHashC")
Private Sub Test03a_Add_MultipleItems()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As cHashC
    Dim myExpected As Variant
    myExpected = Array(42, "Hello", 3.142)
    Sorters.ShakerSortArrayByIndex myExpected
    Dim myresult As Variant

    'Act:
    Set myH = New cHashC
    myH.Add 42
    myH.Add "Hello"
    myH.Add 3.142
    
    myresult = myH.Items
    Sorters.ShakerSortArrayByIndex myresult
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("cHashC")
Private Sub Test04a0_Remove_SingleItem()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As cHashC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    'ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant
    Set myH = New cHashC
    With myH
        .Add Empty
        .Add Empty
        .Add Empty
        .Add 42
        .Add Empty
        .Add Empty
    End With
        
    'Act:
    myH.Remove 42

    myresult = myH.Items

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("cHashC")
Private Sub Test05a_RemoveByIndex_SingleItem()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As cHashC
    Dim myExpected As Variant
    myExpected = Array(1, 100, 2, 43, 5)
    'ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant
    Set myH = New cHashC
    With myH
        .Add 1
        .Add 100
        .Add 2
        .Add 42
        .Add 43
        .Add 5
    End With
        
    'Act:
    myH.RemoveByIndex 3

    myresult = myH.Items

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("cHashC")
Private Sub Test6a_Clear()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As cHashC
    Dim myExpected As Variant
    myExpected = 0&

    Dim myresult As Variant
    
    Set myH = New cHashC
    With myH
        .Add 1
        .Add 100
        .Add 2
        .Add 42
        .Add 43
        .Add 5
    End With

    'Act:
    myH.Clear
    myresult = myH.Count

    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("cHashC")
Private Sub Test7a_Exists_True()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As cHashC
    Dim myExpected As Boolean
    myExpected = True

    Dim myresult As Boolean
    
    Set myH = New cHashC
    With myH
        .Add 1
        .Add 100&
        .Add 2
        .Add 42
        .Add 43
        .Add 5
    End With

    'Act:
    myresult = myH.Exists(100&)

    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("cHashC")
Private Sub Test7b_Exists_False()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As cHashC
    Dim myExpected As Boolean
    myExpected = False

    Dim myresult As Boolean
    
    Set myH = New cHashC
    With myH
        .Add 1
        .Add 100&
        .Add 2
        .Add 42
        .Add 43
        .Add 5
    End With

    'Act:
    myresult = myH.Exists(1000&)

    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

