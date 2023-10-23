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
        Debug.Print CurrentProcedureName ;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
    #End If

    Test01_SeqObj
    Test02_Initialised
    Test03_HeadingAfterRightTurn
    Test04_HeadingAfterLeftTurn
    Test05_MoveSingleStepForward
    Test06_MoveByThreeByTurnRightForEightways
    Test07_MoveByThreeByTurnLeftForEightways
    Test08_MoveByThreeByByHeadingClockwiseForEightways
    Test09_MoveByThreeByByHeadingAntiClockwiseForEightways
    Test10_TrailMoveFiveByNorthEast
    Test10_TrailMoveFiveByNorthEastMovementStats
    
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
    On Error GoTo TestFail
    
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
Private Sub Test03_HeadingAfterRightTurn()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant: myExpected = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 1&, 2&)
    ReDim Preserve myExpected(1 To 10)
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
    
    Dim myS As SeqA: Set myS = SeqA.Deb
   
    Dim myResult As Variant
    
    'Act:
    Dim myCount As Long
    For myCount = 1 To 10
    
        myS.Add myT.Heading
        myT.Turn "R"
    Next
    myResult = myS.ToArray
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

'@TestMethod("TrackedXY")
Private Sub Test04_HeadingAfterLeftTurn()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As Variant: myExpected = Array(1&, 8&, 7&, 6&, 5&, 4&, 3&, 2&, 1&, 8&)
    ReDim Preserve myExpected(1 To 10)
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
    
    Dim myS As SeqA: Set myS = SeqA.Deb
   
    Dim myResult As Variant
    
    'Act:
    Dim myCount As Long
    For myCount = 1 To 10
    
        myS.Add myT.Heading
        myT.Turn "L"
    Next
    myResult = myS.ToArray
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

'@TestMethod("TrackedXY")
Private Sub Test05_MoveSingleStepForward()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String: myExpected = "{{0,0},{0,1}}"
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
   
    Dim myResult As Variant
    
    'Act:
    myT.Move
   
    myResult = Fmt.Text("{0}", myT.Track)
    
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

'@TestMethod("TrackedXY")
Private Sub Test06_MoveByThreeByTurnRightForEightways()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String: myExpected = "{{0,0},{0,3},{3,6},{6,6},{9,3},{9,0},{6,-3},{3,-3},{0,0}}"
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
   
    Dim myResult As String
    Dim myS As SeqA: Set myS = SeqA.Deb
    'Act:
    
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "R"
    myS.Add myT.Location
    
   
    myResult = Fmt.Text("{0}", myS)
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



'@TestMethod("TrackedXY")
Private Sub Test07_MoveByThreeByTurnLeftForEightways()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String: myExpected = "{{0,0},{0,3},{-3,6},{-6,6},{-9,3},{-9,0},{-6,-3},{-3,-3},{0,0}}"
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
   
    Dim myResult As String
    Dim myS As SeqA: Set myS = SeqA.Deb
    'Act:
    
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    myT.Move 3
    myT.Turn "L"
    myS.Add myT.Location
    
   
    myResult = Fmt.Text("{0}", myS)
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

'@TestMethod("TrackedXY")
Private Sub Test08_MoveByThreeByByHeadingClockwiseForEightways()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String: myExpected = "{{0,0},{0,3},{3,6},{6,6},{9,3},{9,0},{6,-3},{3,-3},{0,0}}"
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
   
    Dim myResult As String
    Dim myS As SeqA: Set myS = SeqA.Deb
    'Act:
    
    myS.Add myT.Location
    myT.Move 3, "North"
    myS.Add myT.Location
    myT.Move 3, "NE"
    myS.Add myT.Location
    myT.Move 3, "East"
    myS.Add myT.Location
    myT.Move 3, "SE"
    myS.Add myT.Location
    myT.Move 3, "South"
    myS.Add myT.Location
    myT.Move 3, "SW"
    myS.Add myT.Location
    myT.Move 3, "West"
    myS.Add myT.Location
    myT.Move 3, "NorthWest"
    myS.Add myT.Location
    
    myResult = Fmt.Text("{0}", myS)
   
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

'@TestMethod("TrackedXY")
Private Sub Test09_MoveByThreeByByHeadingAntiClockwiseForEightways()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:

    Dim myExpected As String: myExpected = "{{0,0},{0,3},{-3,6},{-6,6},{-9,3},{-9,0},{-6,-3},{-3,-3},{0,0}}"
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
   
    Dim myResult As String
    Dim myS As SeqA: Set myS = SeqA.Deb
    'Act:
    
    myS.Add myT.Location
    myT.Move 3, "North"
    myS.Add myT.Location
    myT.Move 3, "NW"
    myS.Add myT.Location
    myT.Move 3, "West"
    myS.Add myT.Location
    myT.Move 3, "SW"
    myS.Add myT.Location
    myT.Move 3, "South"
    myS.Add myT.Location
    myT.Move 3, "SE"
    myS.Add myT.Location
    myT.Move 3, "East"
    myS.Add myT.Location
    myT.Move 3, "NorthEast"
    myS.Add myT.Location
    
   
    myResult = Fmt.Text("{0}", myS)
  
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

'@TestMethod("TrackedXY")
Private Sub Test10_TrailMoveFiveByNorthEast()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:

    Dim myExpected As String: myExpected = "{{0,0},{1,1},{2,2},{3,3},{4,4},{5,5}}"
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
   
    Dim myResult As String
    
    'Act:
    
   myT.Move(5, "NE")
    myResult = Fmt.Text("{0}", myT.Track)
  
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

'@TestMethod("TrackedXY")
Private Sub Test10_TrailMoveFiveByNorthEastMovementStats()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:

    Dim myExpected As Variant: myExpected = Array(True, 5&, e_Heading.m_NE, False, False, 10, "5,5")
    Dim myT As TrackedXY
    Set myT = TrackedXY.Deb
   
    Dim myResult(0 To 6) As Variant
    
    
    'Act:
    
    myT.Move(5, "NE")
    myResult(0) = myT.HasTurned
    myResult(1) = myT.StepsTaken
    myResult(2) = myT.Heading
    myResult(3) = myT.BoundsInUse
    myResult(4) = myT.AtOrigin
    myResult(5) = myT.Manhatten
    myResult(6) = myT.Location.ToString
  
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