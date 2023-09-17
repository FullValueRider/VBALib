Attribute VB_Name = "TestArrayOp"
'@IgnoreModule
'@TestModule
'@Folder("Tests")
'@PrivateModule
Option Explicit
Option Private Module

'Public Assert As Object
'Public Fakes As Object

#If TWINBASIC Then
    'Do nothing
#Else


    '@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    GlobalAssert
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

Public Sub ArrayOpTests()

    #If TWINBASIC Then
        Debug.Print CurrentProcedureName; vbTab, vbTab,
    #Else
        GlobalAssert
        VBATesting = True
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
    VBATesting = False
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


'@TestMethod("ArrayOp")
Public Sub Test01a_HoldsItems()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
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
    Dim myresult(0 To 7) As Variant
    myresult(0) = CVar(ArrayOp.HoldsItems(myLong))
    myresult(1) = CVar(ArrayOp.HoldsItems(myvar))
    myresult(2) = CVar(ArrayOp.HoldsItems(myArray1))
    myresult(3) = CVar(ArrayOp.HoldsItems(myArray2))
    myresult(4) = CVar(ArrayOp.HoldsItems(myArray3))
    myresult(5) = CVar(ArrayOp.HoldsItems(myArray4))
    
    myresult(6) = CVar(ArrayOp.HoldsItems(myArray5))
    myresult(7) = CVar(ArrayOp.HoldsItems(myArray6))
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayOp")
Public Sub Test01b_LacksItems()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
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
    Dim myresult(0 To 7) As Variant
    myresult(0) = CVar(ArrayOp.LacksItems(myLong))
    myresult(1) = CVar(ArrayOp.LacksItems(myvar))
    myresult(2) = CVar(ArrayOp.LacksItems(myArray1))
    myresult(3) = CVar(ArrayOp.LacksItems(myArray2))
    myresult(4) = CVar(ArrayOp.LacksItems(myArray3))
    myresult(5) = CVar(ArrayOp.LacksItems(myArray4))
    
    myresult(6) = CVar(ArrayOp.LacksItems(myArray5))
    myresult(7) = CVar(ArrayOp.LacksItems(myArray6))
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayOp")
Public Sub Test02a_IsArray()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
    'Arrange
    Dim myExpected As Variant: myExpected = Array(False, True, True, True, True, True, True, False, False, False)

    Dim myresult As Variant
    ReDim myresult(0 To 9)
    'act
    Dim myEMptyArray As Variant: myEMptyArray = Array()
    Dim myListArray As Variant: myListArray = Array(1, 2, 3, 4, 5)
    Dim myTableArray As Variant: myTableArray = MakeTableArray(3, 3)
    Dim my3dArray As Variant: my3dArray = Make3DArray(3, 3, 3)
    
    myresult(0) = ArrayOp.IsArray(myEMptyArray)
    myresult(1) = ArrayOp.IsArray(myListArray)
    myresult(2) = ArrayOp.IsArray(myTableArray)
    myresult(3) = ArrayOp.IsArray(my3dArray)
    
    myresult(4) = ArrayOp.IsArray(myListArray, m_ListArray)
    myresult(5) = ArrayOp.IsArray(myTableArray, m_TableArray)
    myresult(6) = ArrayOp.IsArray(my3dArray, m_MDArray)
    
    myresult(7) = ArrayOp.IsArray(myListArray, m_TableArray)
    myresult(8) = ArrayOp.IsArray(myTableArray, m_MDArray)
    myresult(9) = ArrayOp.IsArray(my3dArray, m_ListArray)
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayOp")
Public Sub Test02b_IsNotArray()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
    'Arrange
    Dim myExpected As Variant: myExpected = Array(True, False, False, False, False, False, False, True, True, True)
                                            
    Dim myresult As Variant
    ReDim myresult(0 To 9)
    'act
    Dim myEMptyArray As Variant: myEMptyArray = Array()
    Dim myListArray As Variant: myListArray = Array(1, 2, 3, 4, 5)
    Dim myTableArray As Variant: myTableArray = MakeTableArray(3, 3)
    Dim my3dArray As Variant: my3dArray = Make3DArray(3, 3, 3)
    
    myresult(0) = ArrayOp.IsNotArray(myEMptyArray)
    myresult(1) = ArrayOp.IsNotArray(myListArray)
    myresult(2) = ArrayOp.IsNotArray(myTableArray)
    myresult(3) = ArrayOp.IsNotArray(my3dArray)
    
    myresult(4) = ArrayOp.IsNotArray(myListArray, m_ListArray)
    myresult(5) = ArrayOp.IsNotArray(myTableArray, m_TableArray)
    myresult(6) = ArrayOp.IsNotArray(my3dArray, m_MDArray)
    
    myresult(7) = ArrayOp.IsNotArray(myListArray, m_TableArray)
    myresult(8) = ArrayOp.IsNotArray(myTableArray, m_MDArray)
    myresult(9) = ArrayOp.IsNotArray(my3dArray, m_ListArray)
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayOp")
Public Sub Test03a_Count()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(-1&, -1&, -1&, 6&, -1&, 5&, 20&, 105&)
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
    Dim myresult(0 To 7) As Variant
    myresult(0) = CVar(ArrayOp.Count(myLong))
    myresult(1) = CVar(ArrayOp.Count(myvar))
    myresult(2) = CVar(ArrayOp.Count(myArray1))
    myresult(3) = CVar(ArrayOp.Count(myArray2))
    myresult(4) = CVar(ArrayOp.Count(myArray3))
    myresult(5) = CVar(ArrayOp.Count(myArray4))
    
    myresult(6) = CVar(ArrayOp.Count(myArray5))
    myresult(7) = CVar(ArrayOp.Count(myArray6))
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayOp")
Public Sub Test04a_Ranks()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(-1&, -1&, 0&, 1&, 0&, 1&, 2&, 3&)
    
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
    Dim myresult(0 To 7) As Variant
    myresult(0) = CVar(ArrayOp.Ranks(myLong))    ' -1
    myresult(1) = CVar(ArrayOp.Ranks(myvar))     ' -1
    myresult(2) = CVar(ArrayOp.Ranks(myArray1))  ' 0
    myresult(3) = CVar(ArrayOp.Ranks(myArray2))  ' 1
    myresult(4) = CVar(ArrayOp.Ranks(myArray3))  ' 0
    myresult(5) = CVar(ArrayOp.Ranks(myArray4))  ' 1
    myresult(6) = CVar(ArrayOp.Ranks(myArray5))  ' 2
    myresult(7) = CVar(ArrayOp.Ranks(myArray6))  ' 3
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


''@TestMethod("ArrayOp")
Public Sub Test04b_HoldsRank()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
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
    Dim myresult(0 To 7) As Variant
    myresult(0) = CVar(ArrayOp.HoldsRank(myLong, 2))
    myresult(1) = CVar(ArrayOp.HoldsRank(myvar, 2))
    myresult(2) = CVar(ArrayOp.HoldsRank(myArray1, 2))
    myresult(3) = CVar(ArrayOp.HoldsRank(myArray2, 2))
    myresult(4) = CVar(ArrayOp.HoldsRank(myArray3, 2))
    myresult(5) = CVar(ArrayOp.HoldsRank(myArray4, 2))

    myresult(6) = CVar(ArrayOp.HoldsRank(myArray5, 2))
    myresult(7) = CVar(ArrayOp.HoldsRank(myArray6, 2))

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName

TestExit:
    Exit Sub

TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub


'@TestMethod("ArrayOp")
Public Sub Test04c_LacksRank()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
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
    Dim myresult(0 To 7) As Variant
    myresult(0) = CVar(ArrayOp.HoldsRank(myLong, 2))
    myresult(1) = CVar(ArrayOp.HoldsRank(myvar, 2))
    myresult(2) = CVar(ArrayOp.HoldsRank(myArray1, 2))
    myresult(3) = CVar(ArrayOp.HoldsRank(myArray2, 2))
    myresult(4) = CVar(ArrayOp.HoldsRank(myArray3, 2))
    myresult(5) = CVar(ArrayOp.HoldsRank(myArray4, 2))

    myresult(6) = CVar(ArrayOp.HoldsRank(myArray5, 2))
    myresult(7) = CVar(ArrayOp.HoldsRank(myArray6, 2))

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName

TestExit:
    Exit Sub

TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit

End Sub


'@TestMethod("ArrayOp")
Public Sub Test05a_FirstIndex()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(Null, Null, Null, 0&, Null, 1&, 1&, 3&)
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
    Dim myresult As Variant
    ReDim myresult(0 To 7)
    myresult(0) = CVar(ArrayOp.FirstIndex(myLong))
    myresult(1) = CVar(ArrayOp.FirstIndex(myvar))
    myresult(2) = CVar(ArrayOp.FirstIndex(myArray1))
    myresult(3) = CVar(ArrayOp.FirstIndex(myArray2))
    myresult(4) = CVar(ArrayOp.FirstIndex(myArray3))
    myresult(5) = CVar(ArrayOp.FirstIndex(myArray4, 1))
    
    myresult(6) = CVar(ArrayOp.FirstIndex(myArray5, 2))
    myresult(7) = CVar(ArrayOp.FirstIndex(myArray6, 3))
    
    
    myExpected = ArrayOp.MapIt(myExpected, mpReplaceNull.Deb)
    myresult = ArrayOp.MapIt(myresult, mpReplaceNull.Deb)
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
  
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("ArrayOp")
Public Sub Test05b_LastIndex()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
   
    'Arrange:
    Dim myExpected  As Variant: myExpected = Array(Null, Null, Null, 5&, Null, 5&, 5&, 9&)
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
    Dim myresult(0 To 7) As Variant
    myresult(0) = CVar(ArrayOp.Lastindex(myLong))
    myresult(1) = CVar(ArrayOp.Lastindex(myvar))
    myresult(2) = CVar(ArrayOp.Lastindex(myArray1))
    myresult(3) = CVar(ArrayOp.Lastindex(myArray2))
    myresult(4) = CVar(ArrayOp.Lastindex(myArray3))
    myresult(5) = CVar(ArrayOp.Lastindex(myArray4, 1))
    
    myresult(6) = CVar(ArrayOp.Lastindex(myArray5, 2))
    myresult(7) = CVar(ArrayOp.Lastindex(myArray6, 3))
    
    'Assert:
    AssertExactSequenceEquals ArrayOp.MapIt(myExpected, mpReplaceNull.Deb), ArrayOp.MapIt(myresult, mpReplaceNull.Deb), myProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


