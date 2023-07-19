Attribute VB_Name = "TestArrayInfo"
'@IgnoreModule
'@TestModule
'@Folder("Tests")
'@PrivateModule
Option Explicit
Option Private Module

'Public Assert As Object
'Public Fakes As Object

#If twinbasic Then
    'Do nothing
#Else

    '@ModuleInitialize
    Public Sub ModuleInitialize()
        'this method runs once per module.
    End Sub
    
    '@ModuleCleanup
    Public Sub ModuleCleanup()
        'this method runs once per module.
    End Sub
    
    '@TestInitialize
    Public Sub TestInitialize()
        'This method runs before every test in the module..
    End Sub
    
    '@TestCleanup
    Public Sub TestCleanup()
        'this method runs after every test in the module.
    End Sub
    
#End If

Public Sub ArrayInfoTests()

#If twinbasic Then
    Debug.Print CurrentProcedureName; vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    Test01a_HoldsItems
    Test01b_LacksItems
    
    Test02a_IsArray
    Test02b_IsNotArray
    
    Test03a_Count
    
    Test04a_Ranks
    Test04b_HoldsRank
    Test04c_LacksRank
    
    Test05a_FirstIndex
    Test05b_LastIndex
    
    Debug.Print "Testing completed"

End Sub
    

Public Function MakeTableArray(ByVal ipFirst As Long, ByVal ipSecond As Long) As Variant

    '@Ignore VariableNotAssigned
    Dim myArray As Variant
    ReDim myArray(1 To ipFirst, 1 To ipSecond)
    Dim myValue As Long
    myValue = 1
    
    Dim myFirst As Long
    For myFirst = 1 To ipFirst
    
        Dim mySecond As Long
        For mySecond = 1 To ipSecond
        
            myArray(myFirst, mySecond) = myValue
            myValue = myValue + 1
            
        Next
        
    Next
        
    MakeTableArray = myArray
    
End Function


Public Function Make3DArray(ByVal ipFirst As Long, ByVal ipSecond As Long, ByVal ipThird As Variant) As Variant

    '@Ignore VariableNotAssigned
    Dim myArray As Variant
    ReDim myArray(1 To ipFirst, 1 To ipSecond, 1 To ipThird)
    Dim myValue As Long
    myValue = 1
    
    Dim myFirst As Long
    For myFirst = 1 To ipFirst
    
        Dim mySecond As Long
        For mySecond = 1 To ipSecond
        
            Dim myThird As Long
            For myThird = 1 To ipThird
                myArray(myFirst, mySecond, myThird) = myValue
                myValue = myValue + 1
            Next
        Next
    Next
        
    Make3DArray = myArray
    
End Function

Public Function GetParamArray(ParamArray ipArgs() As Variant) As Variant
    GetParamArray = ipArgs
End Function

'@TestMethod("ArrayInfo")
Public Sub Test01a_HoldsItems()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(False, False, False, True, False, True, True, True)
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String
    
    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.HoldsItems(myLong))
    myResult(1) = CVar(ArrayInfo.HoldsItems(myvar))
    myResult(2) = CVar(ArrayInfo.HoldsItems(myArray1))
    myResult(3) = CVar(ArrayInfo.HoldsItems(myArray2))
    myResult(4) = CVar(ArrayInfo.HoldsItems(myArray3))
    myResult(5) = CVar(ArrayInfo.HoldsItems(myArray4))
    
    myResult(6) = CVar(ArrayInfo.HoldsItems(myArray5))
    myResult(7) = CVar(ArrayInfo.HoldsItems(myArray6))
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("ArrayInfo")
Public Sub Test01b_LacksItems()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(True, True, True, False, True, False, False, False)
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String
    
    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.LacksItems(myLong))
    myResult(1) = CVar(ArrayInfo.LacksItems(myvar))
    myResult(2) = CVar(ArrayInfo.LacksItems(myArray1))
    myResult(3) = CVar(ArrayInfo.LacksItems(myArray2))
    myResult(4) = CVar(ArrayInfo.LacksItems(myArray3))
    myResult(5) = CVar(ArrayInfo.LacksItems(myArray4))
    
    myResult(6) = CVar(ArrayInfo.LacksItems(myArray5))
    myResult(7) = CVar(ArrayInfo.LacksItems(myArray6))
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayInfo")
Public Sub Test02a_IsArray()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
    'Arrange
    Dim myExpected As Variant: myExpected = Array(False, True, True, True, True, True, True, False, False, False)

    Dim myResult As Variant
    ReDim myResult(0 To 9)
    'act
    Dim myEMptyArray As Variant: myEMptyArray = Array()
    Dim myListArray As Variant: myListArray = Array(1, 2, 3, 4, 5)
    Dim myTableArray As Variant: myTableArray = MakeTableArray(3, 3)
    Dim my3dArray As Variant: my3dArray = Make3DArray(3, 3, 3)
    
    myResult(0) = True
    myResult(1) = ArrayInfo.IsArray(myListArray)
    myResult(2) = ArrayInfo.IsArray(myTableArray)
    myResult(3) = ArrayInfo.IsArray(my3dArray)
    
    myResult(1) = ArrayInfo.IsArray(myListArray, m_ListArray)
    myResult(2) = ArrayInfo.IsArray(myTableArray, m_TableArray)
    myResult(3) = ArrayInfo.IsArray(my3dArray, m_MDArray)
    
    myResult(1) = ArrayInfo.IsArray(myListArray, m_TableArray)
    myResult(2) = ArrayInfo.IsArray(myTableArray, m_MDArray)
    myResult(3) = ArrayInfo.IsArray(my3dArray, m_ListArray)
    'Assert:
    Assert.SequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("ArrayInfo")
Public Sub Test02b_IsNotArray()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
    'Arrange
    Dim myExpected As Variant: myExpected = Array(True, False, False, False, False, False, False, True, True, True)
                                            
    Dim myResult As Variant
    ReDim myResult(0 To 3)
    'act
    Dim myEMptyArray As Variant: myEMptyArray = Array()
    Dim myListArray As Variant: myListArray = Array(1, 2, 3, 4, 5)
    Dim myTableArray As Variant: myTableArray = MakeTableArray(3, 3)
    Dim my3dArray As Variant: my3dArray = Make3DArray(3, 3, 3)
    
    myResult(0) = ArrayInfo.IsArray(myEMptyArray)
    myResult(1) = ArrayInfo.IsArray(myListArray)
    myResult(2) = ArrayInfo.IsArray(myTableArray)
    myResult(3) = ArrayInfo.IsArray(my3dArray)
    
    myResult(1) = ArrayInfo.IsArray(myListArray, m_ListArray)
    myResult(2) = ArrayInfo.IsArray(myTableArray, m_TableArray)
    myResult(3) = ArrayInfo.IsArray(my3dArray, m_MDArray)
    
    myResult(1) = ArrayInfo.IsArray(myListArray, m_TableArray)
    myResult(2) = ArrayInfo.IsArray(myTableArray, m_MDArray)
    myResult(3) = ArrayInfo.IsArray(my3dArray, m_ListArray)
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayInfo")
Public Sub Test03a_Count()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(-1, -1, -1, 6, -1, 5, 20, 105)
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String
    
    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.Count(myLong))
    myResult(1) = CVar(ArrayInfo.Count(myvar))
    myResult(2) = CVar(ArrayInfo.Count(myArray1))
    myResult(3) = CVar(ArrayInfo.Count(myArray2))
    myResult(4) = CVar(ArrayInfo.Count(myArray3))
    myResult(5) = CVar(ArrayInfo.Count(myArray4))
    
    myResult(6) = CVar(ArrayInfo.Count(myArray5))
    myResult(7) = CVar(ArrayInfo.Count(myArray6))
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayInfo")
Public Sub Test04a_Ranks()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(-1, -1, 0, 1, 0, 1, 2, 3)
    
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String
    
    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.Ranks(myLong))    ' -1
    myResult(1) = CVar(ArrayInfo.Ranks(myvar))     ' -1
    myResult(2) = CVar(ArrayInfo.Ranks(myArray1))  ' 0
    myResult(3) = CVar(ArrayInfo.Ranks(myArray2))  ' 1
    myResult(4) = CVar(ArrayInfo.Ranks(myArray3))  ' 0
    myResult(5) = CVar(ArrayInfo.Ranks(myArray4))  ' 1
    myResult(6) = CVar(ArrayInfo.Ranks(myArray5))  ' 2
    myResult(7) = CVar(ArrayInfo.Ranks(myArray6))  ' 3
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


''@TestMethod("ArrayInfo")
Public Sub Test04b_HoldsRank()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(False, False, False, False, False, False, True, True)
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String

    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.HoldsRank(myLong, 2))
    myResult(1) = CVar(ArrayInfo.HoldsRank(myvar, 2))
    myResult(2) = CVar(ArrayInfo.HoldsRank(myArray1, 2))
    myResult(3) = CVar(ArrayInfo.HoldsRank(myArray2, 2))
    myResult(4) = CVar(ArrayInfo.HoldsRank(myArray3, 2))
    myResult(5) = CVar(ArrayInfo.HoldsRank(myArray4, 2))

    myResult(6) = CVar(ArrayInfo.HoldsRank(myArray5, 2))
    myResult(7) = CVar(ArrayInfo.HoldsRank(myArray6, 2))

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub

TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("ArrayInfo")
Public Sub Test04c_LacksRank()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(False, False, False, False, False, False, True, True)
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String

    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.HoldsRank(myLong, 2))
    myResult(1) = CVar(ArrayInfo.HoldsRank(myvar, 2))
    myResult(2) = CVar(ArrayInfo.HoldsRank(myArray1, 2))
    myResult(3) = CVar(ArrayInfo.HoldsRank(myArray2, 2))
    myResult(4) = CVar(ArrayInfo.HoldsRank(myArray3, 2))
    myResult(5) = CVar(ArrayInfo.HoldsRank(myArray4, 2))

    myResult(6) = CVar(ArrayInfo.HoldsRank(myArray5, 2))
    myResult(7) = CVar(ArrayInfo.HoldsRank(myArray6, 2))

    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    Exit Sub

TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub

'@TestMethod("ArrayInfo")
Public Sub Test05a_FirstIndex()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(Empty, Empty, Empty, 0, Empty, 1, 1, 3)
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String
    
    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.FirstIndex(myLong))
    myResult(1) = CVar(ArrayInfo.FirstIndex(myvar))
    myResult(2) = CVar(ArrayInfo.FirstIndex(myArray1))
    myResult(3) = CVar(ArrayInfo.FirstIndex(myArray2))
    myResult(4) = CVar(ArrayInfo.FirstIndex(myArray3))
    myResult(5) = CVar(ArrayInfo.FirstIndex(myArray4, 1))
    
    myResult(6) = CVar(ArrayInfo.FirstIndex(myArray5, 2))
    myResult(7) = CVar(ArrayInfo.FirstIndex(myArray6, 3))
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("ArrayInfo")
Public Sub Test05b_LastIndex()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(Empty, Empty, Empty, 5, Empty, 5, 5, 9)
    Dim myLong As Long: myLong = 0
    Dim myvar As Variant
    Dim myArray1() As Long
    Dim myArray2(0 To 5) As Long
    Dim myArray3 As Variant: myArray3 = Array()
    Dim myArray4 As Variant
    ReDim myArray4(1 To 5)
    Dim myArray5(1 To 4, 1 To 5) As Long
    Dim myArray6(1 To 5, 2 To 4, 3 To 9) As String
    
    'Act:
    Dim myResult(0 To 7) As Variant
    myResult(0) = CVar(ArrayInfo.LastIndex(myLong))
    myResult(1) = CVar(ArrayInfo.LastIndex(myvar))
    myResult(2) = CVar(ArrayInfo.LastIndex(myArray1))
    myResult(3) = CVar(ArrayInfo.LastIndex(myArray2))
    myResult(4) = CVar(ArrayInfo.LastIndex(myArray3))
    myResult(5) = CVar(ArrayInfo.LastIndex(myArray4, 1))
    
    myResult(6) = CVar(ArrayInfo.LastIndex(myArray5, 2))
    myResult(7) = CVar(ArrayInfo.LastIndex(myArray6, 3))
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

