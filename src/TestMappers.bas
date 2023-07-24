Attribute VB_Name = "TestMappers"
'@TestModule
'@Folder("Tests")
'@IgnoreModule
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

Public Sub MapperTests()
 
#If twinbasic Then
    Debug.Print CurrentProcedureName;
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName;
#End If

    Test01a_mpDec_Default
    Test01b_mpDec_1
    Test01c_mpDec_3
    
    Test02a_mpInc_Default
    Test02b_mpInc_1
    Test02c_mpInc_3
    
    Test03a_mpIndex_mpInc_SeqC
    Test03b_mpIndex_mpInc_Collection
    Test03c_mpIndex_mpInc_ArrayList
    Test03d_mpIndex_mpInc_Array
    Test03e_mpIndex_mpInc_Dictionary
    Test03f_mpIndex_mpInc_String
    
    Test04a_mpInner
    
    Debug.Print vbTab, vbTab, "Testing completed"
    
End Sub
'@TestMethod("Mapper")
Private Sub Test01a_mpDec_Default()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(2&, "Nan", 3&, 4&, 5&, "Nan")
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    
    Dim mySeq As SeqC
    Set mySeq = SeqC(3&, "4", 4&, 5&, 6&, "Six")
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = mySeq.MapIt(mpDec.Deb).ToArray
   
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

'@TestMethod("Mapper")
Private Sub Test01b_mpDec_1()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(2&, "Nan", 3&, 4&, 5&, "Nan")
    ReDim Preserve myExpected(1 To 6)
    
    Dim mySeq As SeqC
    Set mySeq = SeqC.Deb(3&, "4", 4&, 5&, 6&, "Six")
    
    Dim myResult As Variant
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = mySeq.MapIt(mpDec(1)).ToArray
   
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



'@TestMethod("Mapper")
Private Sub Test01c_mpDec_3()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(0&, "Nan", 1&, 2&, 3&, "Nan")
    ReDim Preserve myExpected(1 To 6)
    
    Dim mySeq As SeqC
    Set mySeq = SeqC.Deb(3&, "4", 4&, 5&, 6&, "Six")
    
    Dim myResult As Variant
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = mySeq.MapIt(mpDec(3)).ToArray
   
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


'@TestMethod("Mapper")
Private Sub Test02a_mpInc_Default()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(4&, "5", 5&, 6&, 7&, "Siy")
    ReDim Preserve myExpected(1 To 6)
    
    Dim mySeq As SeqC
    Set mySeq = SeqC.Deb(3&, "4", 4&, 5&, 6&, "Six")
    
    Dim myResult As Variant
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = mySeq.MapIt(mpInc.Deb).ToArray
   
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

'@TestMethod("Mapper")
Private Sub Test02b_mpInc_1()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(4&, "5", 5&, 6&, 7&, "Siy")
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    

    Set mySeq = SeqC.Deb(3&, "4", 4&, 5&, 6&, "Six")
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = mySeq.MapIt(mpInc(1)).ToArray
   
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

'@TestMethod("Mapper")
Private Sub Test02c_mpInc_3()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(6&, "7", 7&, 8&, 9&, "Sj0")
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    

    Set mySeq = SeqC.Deb(3&, "4", 4&, 5&, 6&, "Six")
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = mySeq.MapIt(mpInc(3)).ToArray
   
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

'@TestMethod("Mapper")
Private Sub Test03a_mpIndex_mpInc_SeqC()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant
    myExpected = Array(2, 3, "4")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    ReDim myResult(1 To 3)
    
    Dim mySeq As SeqC
    Set mySeq = SeqC(SeqC(1, 1, 1), SeqC(2, 2, 2), SeqC(3, "3", 3))
    
    'Act:
    Dim myTmp As Variant
    Dim myInc As mpInc: Set myInc = mpInc(1)
    Dim myIndex As mpByIndex: Set myIndex = mpByIndex(myInc, 2)
    Set myTmp = mySeq.MapIt(myIndex)
    myTmp = myTmp.ToArray
    myResult(1) = myTmp(1).Item(2)
    myResult(2) = myTmp(2).Item(2)
    myResult(3) = myTmp(3).Item(2)
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


'@TestMethod("Mapper")
Private Sub Test03b_mpIndex_mpInc_Collection()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant
    myExpected = Array(2, 3, "4")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    ReDim myResult(1 To 3)
    
    Dim myC1 As Collection
    Set myC1 = New Collection
    myC1.Add 1
    myC1.Add 1
    myC1.Add 1
    
    Dim myC2 As Collection
    Set myC2 = New Collection
    myC2.Add 2
    myC2.Add 2
    myC2.Add 2
    
    Dim myC3 As Collection
    Set myC3 = New Collection
    myC3.Add 3
    myC3.Add "3"
    myC3.Add 3
    
    Dim mySeq As SeqC
    Set mySeq = SeqC(myC1, myC2, myC3)
    
    'Act:
    Dim myTmp As Variant
    Set myTmp = mySeq.MapIt(mpByIndex(mpInc(1), 2))
    
    myResult(1) = myTmp.Item(1).Item(2)
    myResult(2) = myTmp.Item(2).Item(2)
    myResult(3) = myTmp.Item(3).Item(2)
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


'@TestMethod("Mapper")
Private Sub Test03c_mpIndex_mpInc_ArrayList()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant
    myExpected = Array(2, 3, "4")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    ReDim myResult(1 To 3)
    
    Dim myAL1 As ArrayList
    Set myAL1 = New ArrayList
    myAL1.Add 1
    myAL1.Add 1
    myAL1.Add 1
    
    Dim myAL2 As ArrayList
    Set myAL2 = New ArrayList
    myAL2.Add 2
    myAL2.Add 2
    myAL2.Add 2
    
    Dim myAL3 As ArrayList
    Set myAL3 = New ArrayList
    myAL3.Add 3
    myAL3.Add "3"
    myAL3.Add 3
    
    Dim mySeq As SeqC
    Set mySeq = SeqC(myAL1, myAL2, myAL3)
    
    'Act:
    Dim myTmp As Variant
    myTmp = mySeq.MapIt(mpByIndex(mpInc(1), 2)).ToArray
    myResult(1) = myTmp(1).Item(2)
    myResult(2) = myTmp(2).Item(2)
    myResult(3) = myTmp(3).Item(2)
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


'@TestMethod("Mapper")
Private Sub Test03d_mpIndex_mpInc_Array()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant
    myExpected = Array(2, 3, "4")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    ReDim myResult(1 To 3)
    
    Dim mySeq As SeqC
    Set mySeq = SeqC(Array(1, 1, 1), Array(2, 2, 2), Array(3, "3", 3))
    
    'Act:
    Dim myTmp As Variant
    myTmp = mySeq.MapIt(mpByIndex(mpInc(1), 2)).ToArray
    myResult(1) = myTmp(1)(2)
    myResult(2) = myTmp(2)(2)
    myResult(3) = myTmp(3)(2)
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


'@TestMethod("Mapper")
Private Sub Test03e_mpIndex_mpInc_Dictionary()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant
    myExpected = Array(2, 3, "4")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    ReDim myResult(1 To 3)
    
    Dim myD1 As KvpC
    Set myD1 = KvpC.Deb
    myD1.Add "one", 1
    myD1.Add "two", 1
    myD1.Add "three", 1
    
    Dim myD2 As KvpC
    Set myD2 = KvpC.Deb
    myD2.Add "one", 2
    myD2.Add "two", 2
    myD2.Add "three", 2
    
    Dim myD3 As KvpC
    Set myD3 = KvpC.Deb
    myD3.Add "one", 3
    myD3.Add "two", "3"
    myD3.Add "three", 3
    
    Dim mySeq As SeqC
    Set mySeq = SeqC(myD1, myD2, myD3)
    
    'Act:
    Dim myTmp As Variant
    myTmp = mySeq.MapIt(mpByIndex(mpInc(1), "two")).ToArray
    myResult(1) = myTmp(1).Item("two")
    myResult(2) = myTmp(2).Item("two")
    myResult(3) = myTmp(3).Item("two")
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


'@TestMethod("Mapper")
Private Sub Test03f_mpIndex_mpInc_String()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant
    myExpected = Array("Iello", "Uhere", "Xorld")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    
    
    Dim mySeq As SeqC
    Set mySeq = SeqC("Hello", "There", "World")
    
    'Act:
    myResult = mySeq.MapIt(mpByIndex(mpInc(1), 1)).ToArray
    
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

'@TestMethod("Mapper")
Private Sub Test04a_mpInner()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    ReDim myExpected(1 To 3)
    myExpected(1) = SeqC(2, 3, 4).ToArray
    myExpected(2) = SeqC(3, 4, 5).ToArray
    myExpected(3) = SeqC(4, 5, 6).ToArray
    
    
    Dim mySeq As SeqC
    Set mySeq = SeqC(SeqC(1, 2, 3), SeqC(2, 3, 4), SeqC(3, 4, 5))
    
    Dim myResult As Variant
    ReDim myResult(1 To 3)
    Dim myTmp As SeqC
    'Act: Apply the mpInc 'function' to each item of the in the inner SeqC
    Set myTmp = mySeq.MapIt(mpInner(mpInc.Deb))
    myResult(1) = myTmp.Item(1).ToArray
    myResult(2) = myTmp.Item(2).ToArray
    myResult(3) = myTmp.Item(3).ToArray
    
    'Assert:
    AssertStrictSequenceEquals myExpected(1), myResult(1), myProcedureName
    AssertStrictSequenceEquals myExpected(2), myResult(2), myProcedureName
    AssertStrictSequenceEquals myExpected(3), myResult(3), myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

