Attribute VB_Name = "TestGlobalFunctors"
'@TestModule
'@Folder("Tests")
'@IgnoreModule
'@ModuleDescription("Global functors implement both IMapper and IReducer")
Option Explicit
Option Private Module
Option Base 1

'Private Assert As Object
'Private Fakes As Object

#If twinbasic Then
    'Do nothing
#Else

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    '    Set Assert = CreateObject("Rubberduck.AssertClass")
    '    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    GlobalAssert
    'this method runs once per module.
    '    Set Assert = Nothing
    '    Set Fakes = Nothing
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


Public Sub GlobalFunctorTests()
 
    #If twinbasic Then
        Debug.Print CurrentProcedureName;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
    #End If

    Test01a_gfIndex
    
    Debug.Print vbTab, vbTab, "Testing completed"
    
End Sub

'@TestMethod("fnKey")
Private Sub Test01a_gfIndex()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myA() As Variant: myA = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100)
    Dim myS As SeqA: Set myS = SeqA(10, 20, 30, 40, 50, 60, 70, 80, 90, 100)
'    Dim myC As Collection: Set myC = New Collection
'    With myC
'        .Add 10
'        .Add 20
'        .Add 30
'        .Add 40
'        .Add 50
'        .Add 60
'        .Add 70
'        .Add 80
'        .Add 90
'        .Add 100
'    End With
    Dim myK As KvpA: Set myK = KvpA().AddPairs(Array("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "None", "Ten"), Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    Dim myStr As String: myStr = "abcdefghij"
    
    Dim myExpected() As Variant: myExpected = Array(50, 50, "e", 50, 50, 50, 50, "e", 50, 50)
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 9)
    
    Dim myKey As IMapper: Set myKey = gfKey(5)
  
    myResult(0) = myKey.ExecMapper(myA)(0)
    myResult(1) = myKey.ExecMapper(myS)(0)
    'myResult(2) = myKey.ExecMapper(myC)(0)
    myResult(2) = myKey.ExecMapper(myStr)(0)
    myResult(3) = myKey.ExecMapper(myK)(0)
    Set myKey = gfKey("Five")
    myResult(4) = myKey.ExecMapper(myK)(0)
    
    Dim myR As IReducer: Set myR = gfKey(5)
    myResult(5) = myR.ExecReduction(myA)(0)
    myResult(6) = myR.ExecReduction(myS)(0)
    'myResult(2) = myKey.ExecMapper(myC)(0)
    myResult(7) = myR.ExecReduction(myStr)(0)
    myResult(8) = myR.ExecReduction(myK)(0)
    Set myR = gfKey("Five")
    myResult(9) = myR.ExecReduction(myK)(0)
    
    'Act:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:

    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
