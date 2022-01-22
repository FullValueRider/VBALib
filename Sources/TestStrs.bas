Attribute VB_Name = "TestStrs"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New Rubberduck.AssertClass
        Set Fakes = New Rubberduck.FakesProvider
    #End If
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


Public Sub StrsTests()

    myInterim = Timer
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Debug.Print "Testing ", ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab, vbTab,
    T01_Dedup
    T02_TrimmerDefault
    T03_TryStringExtentvbNullstring
    T04_TryStringExtentEmptyString
    T05_TryStringExtentHelloWOrld
    
       Debug.Print "completed ", Timer - myInterim
    
End Sub

'@TestMethod("Dedup")
Private Sub T01_Dedup()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As String
        myExpected = "Hello Worlde"
        
        
        Dim myResult As String
        
        'Act:
        myResult = Strs.Dedup("Heeello Worldee", "e")

        'Assert:
        Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print vbCrLf & ErrEx.LiveCallstack.ModuleName, ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
End Sub

'@TestMethod("TrimmerDefault")
Private Sub T02_TrimmerDefault()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As String
        myExpected = "Hello World"
        
        
        Dim myResult As String
        
        'Act:
        myResult = Strs.Trimmer("   ;;;,;,;Hello World ;,; ;; ,")

        'Assert:
        Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print vbCrLf & ErrEx.LiveCallstack.ModuleName, ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
End Sub


'@TestMethod("GetStringExtent")
Public Sub T03_TryStringExtentvbNullstring()

    On Error GoTo TestFail

    Dim myExpectedStatus As Boolean
    myExpectedStatus = False
    
    Dim myItem As String
    myItem = vbNullString
    
    Dim myResult As Result
    
    Set myResult = Strs.TryExtent(myItem)
    
    Assert.AreEqual myExpectedStatus, Globals.Res.Status, ErrEx.LiveCallstack.ProcedureName
    
    Assert.AreEqual myExpectedStatus, myResult.Status, ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ModuleName, ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

'@TestMethod("GetStringExtent")
Public Sub T04_TryStringExtentEmptyString()

    On Error GoTo TestFail

    Dim myExpectedStatus As Boolean
    myExpectedStatus = False
    
    Dim myItem As String
    myItem = ""
    
    Dim myResult As Result
    Dim myResultStatus As Boolean
    myResultStatus = Types.TryExtent(myItem)
    Assert.AreEqual myExpectedStatus, Globals.Res.Status, ErrEx.LiveCallstack.ProcedureName
    Assert.AreEqual myExpectedStatus, myResultStatus, ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ModuleName, ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

'@TestMethod("GetStringExtent")
Public Sub T05_TryStringExtentHelloWOrld()

    On Error GoTo TestFail

    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedItems As Variant
    myExpectedItems = Array(1, 11, 11)
    
    
    Dim myItem As String
    myItem = "Hello World"
    
    Dim myResult As Result
   
    Set myResult = Strs.TryExtent(myItem)
    
    Assert.AreEqual myExpectedStatus, myResult.Status, ErrEx.LiveCallstack.ProcedureName
    Assert.SequenceEquals myExpectedItems, myResult.Items.ToArray, ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ModuleName, ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub


