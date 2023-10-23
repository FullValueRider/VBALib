Attribute VB_Name = "TestPointXY"
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

Public Sub PointXYTests()

    #If twinbasic Then
        Debug.Print CurrentProcedureName ; vbTab, vbTab,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    #End If
    
    Test01_SeqObj
    Test02_InitialisedOK
    Test03_AggregateOutput
    Test04_Offsets_All_N_Clockwise
    Test05_Offsets_NSEW_N_Clockwise
    Test06_Offsets_Diagonals_N_Clockwise
    Test07_Offsets_All_N_Anticlockwise
    Test08_AdjacentCoords_All_N_Clockwise
    Test09_AdjacentCoords_All_SE_Clockwise
    Test10_IndividualPoints
    
    VBATesting = False
    Debug.Print "Testing completed"

End Sub

'@TestMethod("PointXY")
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
    Dim myP As PointXY
    Set myP = PointXY.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "PointXY", "PointXY")
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myResult(0) = VBA.IsObject(myP)
    myResult(1) = VBA.TypeName(myP)
    myResult(2) = myP.TypeName
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


'@TestMethod("PointXY")
Private Sub Test02_InitialisedOK()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY.Deb
    Dim myExpected As Variant
    myExpected = Array(0&, 0&) ', False, True)
    
    Dim myResult(0 To 1) As Variant
    
    'Act:
    myResult(0) = myP.X
    myResult(1) = myP.Y
'    myresult(2) = myP.BoundsInUse
'    myresult(3) = myP.Forbidden Is Nothing
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

'@TestMethod("PointXY")
Private Sub Test03_AggregateOutput()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY.Deb
    Dim myExpected As Variant
    myExpected = Array("0,0", "0,0", "0,0")
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myResult(0) = myP.ToString
    myResult(1) = Fmt.SetArrayMarkupSeparatorOnly.Text("{0}", Array(myP.ToArray))
    myResult(2) = Fmt.SetObjectMarkup(vbNullString, vbNullString, vbNullString).SetDictionaryItemMarkupSeparatorOnly.Text("{0}", myP.ToKVPair)
    
    
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

'@TestMethod("PointXY")
Private Sub Test04_Offsets_All_N_Clockwise()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY(5, 5)
    Dim myExpected As String
    myExpected = "{{0,1},{1,1},{1,0},{1,-1},{0,-1},{-1,-1},{-1,0},{-1,1}}"
    
    Dim myResult As String
    
    'Act:
    myResult = Fmt.Text("{0}", myP.AdjacentCoords(ipAdjacentType:=m_Relative))
    
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

'@TestMethod("PointXY")
Private Sub Test05_Offsets_NSEW_N_Clockwise()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY(5, 5)
    Dim myExpected As String
    myExpected = "{{0,1},{1,0},{0,-1},{-1,0}}"
    
    Dim myResult As String
    
    'Act:
    myResult = Fmt.Text("{0}", myP.AdjacentCoords(m_4WaysNESW, ipAdjacentType:=m_Relative))
    
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

'@TestMethod("PointXY")
Private Sub Test06_Offsets_Diagonals_N_Clockwise()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY(5, 5)
    Dim myExpected As String
    myExpected = "{{1,1},{1,-1},{-1,-1},{-1,1}}"
    
    Dim myResult As String
    
    'Act:
    myResult = Fmt.Text("{0}", myP.AdjacentCoords(e_AdjacentSet.m_4WaysDiagonal, ipAdjacentType:=m_Relative))
    
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


'@TestMethod("PointXY")
Private Sub Test07_Offsets_All_N_Anticlockwise()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY(5, 5)
    Dim myExpected As String
    myExpected = "{{0,1},{-1,1},{-1,0},{-1,-1},{0,-1},{1,-1},{1,0},{1,1}}"
    
    Dim myResult As String
    
    'Act:
    myResult = Fmt.Text("{0}", myP.AdjacentCoords(ipRotation:=m_Anticlockwise, ipAdjacentType:=m_Relative))
    
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

'@TestMethod("PointXY")
Private Sub Test08_Offsets_All_SE_Clockwise()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY(5, 5)
    Dim myExpected As String
    myExpected = "{{-1,-1},{-1,0},{-1,1},{0,1},{1,1},{1,0},{1,-1},{0,-1}}"
    
    Dim myResult As String
    
    'Act:
    myResult = Fmt.Text("{0}", myP.AdjacentCoords(ipStartDirection:=e_AdjacentDirection.m_SE))
    
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


'@TestMethod("PointXY")
Private Sub Test08_AdjacentCoords_All_N_Clockwise()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY(5, 5)
    Dim myExpected As String
    myExpected = "{{5,6},{6,6},{6,5},{6,4},{5,4},{4,4},{4,5},{4,6}}"
    
    Dim myResult As String
    
    'Act:
    myResult = Fmt.Text("{0}", myP.AdjacentCoords)
    
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

'@TestMethod("PointXY")
Private Sub Test09_AdjacentCoords_All_SE_Clockwise()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myP As PointXY
    Set myP = PointXY(5, 5)
    Dim myExpected As String
    myExpected = "{{6,4},{5,4},{4,4},{4,5},{4,6},{5,6},{6,6},{6,5}}"
    
    Dim myResult As String
    
    'Act:
    myResult = Fmt.Text("{0}", myP.AdjacentCoords(ipStartDirection:=e_AdjacentDirection.m_SE))
    
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

'@TestMethod("PointXY")
Private Sub Test10_IndividualPoints()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myS As SeqA: Set myS = SeqA()
    Dim myP As PointXY: Set myP = PointXY(5, 5)
    
    Dim myDirection As e_AdjacentDirection
    For myDirection = e_AdjacentDirection.m_First To e_AdjacentDirection.m_Last
        myS.Add myP.AdjacentCoord(myDirection)
        myS.Add myP.AdjacentCoord(myDirection, ipAdjacentType:=m_Relative)
    Next
    
    
    Dim myExpected As String
    myExpected = "{{5,6},{0,1},{6,6},{1,1},{6,5},{1,0},{6,4},{1,-1},{5,4},{0,-1},{4,4},{-1,-1},{4,5},{-1,0},{4,6},{-1,1}}"
    
    Dim myResult As String
    
    'Act:
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
