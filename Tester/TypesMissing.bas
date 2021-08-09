Attribute VB_Name = "TypesMissing"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
'Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
'    Set Fakes = New Rubberduck.FakesProvider
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
'    Set Fakes = Nothing
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

Private Function GetOptional(ByVal ipFixed As Variant, Optional ByVal ipOptional As Variant) As Variant
    If VBA.IsMissing(ipOptional) Then
    
        GetOptional = ipOptional
        
    Else
    
        GetOptional = ipFixed
        
    End If
    
End Function

'@TestMethod("Missing")
Private Sub Test01_ParameterPresentIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myOptional As Variant
    'Act:
    myOptional = GetOptional(42)
    myResult = Types.IsNotMissing(myOptional)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Missing")
Private Sub Test02_ParameterMissingIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    Dim myOptional As Variant
    'Act:
    myOptional = GetOptional(42, 42)
    myResult = Types.IsNotMissing(myOptional)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

