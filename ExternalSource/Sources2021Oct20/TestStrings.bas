Attribute VB_Name = "TestStrings"
'@IgnoreModule
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
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
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
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
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
End Sub
