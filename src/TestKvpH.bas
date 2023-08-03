Attribute VB_Name = "TestKvpH"
'@IgnoreModule
'@TestModule
'@Folder("Tests")
Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

#If twinbasic Then
    'Do nothing
#Else


    '@ModuleInitialize
Private Sub ModuleInitialize()
    GlobalAssert
    'this method runs once per module.
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

Public Sub KvpHTests()

    #If twinbasic Then
        Debug.Print CurrentProcedureName; vbTab, vbTab,
    #Else
        GlobalAssert
        Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    #End If

    Test01_IsKvpH
    Test02a_Add_MultipleItems_Count
    Test03a_Add_MultipleItems
    Test04a0_Remove_SingleItem
    Test04bc_Remove_LastItem
    Test04bc_Remove_LastItem
    Test05a_RemoveAt_SingleItem
    Test6a_Clear
    Test7a_Exists_True
    Test7b_Exists_False
    
    Debug.Print vbTab, "Testing completed"

End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test01_IsKvpH()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myH As KvpH
    Set myH = KvpH.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "KvpH", "KvpH")
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myResult(0) = VBA.IsObject(myH)
    myResult(1) = VBA.Typename(myH)
    myResult(2) = myH.Typename
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test02a_Add_MultipleItems_Count()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Variant
    myExpected = 3&

    Dim myResult As Variant

    'Act:
    Set myH = KvpH.Deb
    myH.Add 42, "Hello"
    myH.Add "Hello", "There"
    myH.Add 3.142, "World"

    myResult = myH.Count
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test03a_Add_MultipleItems()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Variant
    myExpected = Array(42, "Hello", 3.142)
    ReDim Preserve myExpected(1 To 3)
    Sorters.ShakerSortArray myExpected
    Dim myResult As Variant

    'Act:
    Set myH = KvpH.Deb.ForbidSameKeys
    myH.Add 42, "Hello"
    myH.Add "Hello", "There"
    myH.Add 3.142, "World"
    
    myResult = myH.Keys
    Sorters.ShakerSortArray myResult
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test04a0_Remove_SingleItem()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Variant
    myExpected = Array("One", "Two", "Four", "Five", "Six")
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set myH = KvpH.Deb
    With myH
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
        .Add "Five", 50
        .Add "Six", 60
    End With

    'Act:
    myH.Remove "Three"

    myResult = myH.Keys

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test04b0_Remove_FirstItem()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Variant
    myExpected = Array("Two", "Three", "Four", "Five", "Six")
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set myH = KvpH.Deb
    With myH
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
        .Add "Five", 50
        .Add "Six", 60
    End With

    'Act:
    myH.Remove "One"

    myResult = myH.Keys

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test04bc_Remove_LastItem()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Variant
    myExpected = Array("One", "Two", "Three", "Four", "Five")
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set myH = KvpH.Deb
    With myH
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
        .Add "Five", 50
        .Add "Six", 60
    End With

    'Act:
    myH.Remove "Six"

    myResult = myH.Keys

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test05a_RemoveAt_SingleItem()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Variant
    myExpected = Array("One", "Two", "Four", "Five", "Six")
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set myH = KvpH.Deb
    With myH
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
        .Add "Five", 50
        .Add "Six", 60
    End With

    'Act:
    myH.RemoveAt 3

    myResult = myH.Keys

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test6a_Clear()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Variant
    myExpected = -1

    Dim myResult As Variant

    Set myH = KvpH.Deb
    With myH
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
        .Add "Five", 50
        .Add "Six", 60
    End With


    'Act:
    myH.Clear
    myResult = myH.Count

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test7a_Exists_True()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Boolean
    myExpected = True

    Dim myResult As Boolean

    Set myH = KvpH.Deb
    With myH
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
        .Add "Five", 50
        .Add "Six", 60
    End With

    'Act:
    myResult = myH.Exists("Three")

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test7b_Exists_False()
    On Error GoTo TestFail

    'Arrange:
    Dim myH As KvpH
    Dim myExpected As Boolean
    myExpected = False

    Dim myResult As Boolean

    Set myH = KvpH.Deb
    With myH
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
        .Add "Five", 50
        .Add "Six", 60
    End With

    'Act:
    myResult = myH.Exists("FortyTwo")

    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


