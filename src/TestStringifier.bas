Attribute VB_Name = "TestStringifier"
'@TestModule
'@Folder("Tests")
'@IgnoreModule
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


Public Sub StringifierTests()
 
    #If twinbasic Then
        Debug.Print CurrentProcedureName, vbTab,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName, vbTab,
    #End If

    Test01_StringifyItem_String
    Test01a_StringifyItemWIthTypes_String
    Test01b_StringifyItemWIthTypesInner_String
    Test01c_StringifyItemWIthTypesOuter_String
    
    Test02_StringifyItem_Long
    Test02a_StringifyItemWithTypes_Long
    Test02b_StringifyItemWithTypesInner_Long
    Test02c_StringifyItemWithTypesOuter_Long
    
    Test03_StringifyItem_Array
    Test03a_StringifyItemWithTypes_Array
    Test03b_StringifyItemWithTypesInner_Array
    Test03c_StringifyItemWithTypesOuter_Array
    
    Test04_StringifyItem_SeqC
    Test04a_StringifyItemWithTypes_SeqC
    Test04b_StringifyItemWithTypesInner_SeqC
    Test04c_StringifyItemWithTypesOuter_SeqC
    
    Test05_StringifyItem_Collection
    Test05a_StringifyItemWithTypes_Collection
    Test05b_StringifyItemWithTypesInner_Collection
    Test05c_StringifyItemWithTypesOuter_Collection
    
    Test06_StringifyItem_Dictionary
    Test06a_StringifyItemWithTypes_Dictionary
    Test06b_StringifyItemWithTypesInner_Dictionary
    Test06c_StringifyItemWithTypesOuter_Dictionary
    
    Test07_StringifyItem_CustomDictionaryMarkup
    Test07a_StringifyItemWithTypes_CustomDictionaryMarkup
    Test07b_StringifyItemWithTypesInner_CustomDictionaryMarkup
    Test07c_StringifyItemWithTypesOuter_CustomDictionaryMarkup
    
    Debug.Print "Testing completed"

End Sub


'@TestMethod("Stringifier")
Private Sub Test01_StringifyItem_String()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World!"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier.Deb.StringifyItem("Hello World!")
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test01a_StringifyItemWIthTypes_String()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "String: Hello World!"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier.Deb.WithTypes(e_WithTypes.m_All).StringifyItem("Hello World!")
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test01b_StringifyItemWIthTypesInner_String()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "String: Hello World!"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier(e_WithTypes.m_Inner).StringifyItem("Hello World!")
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test01c_StringifyItemWIthTypesOuter_String()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World!"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier(e_WithTypes.m_Outer).StringifyItem("Hello World!")
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test02_StringifyItem_Long()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "42"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier.Deb.ResetMarkup.StringifyItem(42)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test02a_StringifyItemWithTypes_Long()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Integer: 42"  ' Yes, its an integer, not Long
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier.Deb.WithTypes(e_WithTypes.m_All).ResetMarkup.StringifyItem(42)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test02b_StringifyItemWithTypesInner_Long()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Integer: 42"  ' Yes, its an integer, not Long
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier(e_WithTypes.m_Inner).ResetMarkup.StringifyItem(42)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Stringifier")
Private Sub Test02c_StringifyItemWithTypesOuter_Long()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "42"  ' Yes, its an integer, not Long
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier(e_WithTypes.m_Outer).ResetMarkup.StringifyItem(42)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test03_StringifyItem_Array()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "[1,2,3,4,5,6]"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier.Deb.ResetMarkup.StringifyItem(Array(1, 2, 3, 4, 5, 6))
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test03a_StringifyItemWithTypes_Array()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Variant(): [Integer: 1,Integer: 2,Integer: 3,Integer: 4,Integer: 5,Integer: 6]"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier.Deb.WithTypes(e_WithTypes.m_All).ResetMarkup.StringifyItem(Array(1, 2, 3, 4, 5, 6))
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test03b_StringifyItemWithTypesInner_Array()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "[Integer: 1,Integer: 2,Integer: 3,Integer: 4,Integer: 5,Integer: 6]"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier(e_WithTypes.m_Inner).ResetMarkup.StringifyItem(Array(1, 2, 3, 4, 5, 6))
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Stringifier")
Private Sub Test03c_StringifyItemWithTypesOuter_Array()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Variant(): [1,2,3,4,5,6]"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Stringifier(e_WithTypes.m_Outer).ResetMarkup.StringifyItem(Array(1, 2, 3, 4, 5, 6))
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test04_StringifyItem_SeqC()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{1,2,3,4,5,6}"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Stringifier.ResetMarkup
    Dim myC As SeqC: Set myC = SeqC(1, 2, 3, 4, 5, 6)
    myResult = Stringifier.Deb.ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test04a_StringifyItemWithTypes_SeqC()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "SeqC: {Integer: 1,Integer: 2,Integer: 3,Integer: 4,Integer: 5,Integer: 6}"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Stringifier.ResetMarkup
    Dim myC As SeqC: Set myC = SeqC(1, 2, 3, 4, 5, 6)
    myResult = Stringifier.Deb.WithTypes((e_WithTypes.m_All)).ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test04b_StringifyItemWithTypesInner_SeqC()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{Integer: 1,Integer: 2,Integer: 3,Integer: 4,Integer: 5,Integer: 6}"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Stringifier.ResetMarkup
    Dim myC As SeqC: Set myC = SeqC(1, 2, 3, 4, 5, 6)
    myResult = Stringifier(e_WithTypes.m_Inner).ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test04c_StringifyItemWithTypesOuter_SeqC()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "SeqC: {1,2,3,4,5,6}"
    
    Dim myResult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Stringifier.ResetMarkup
    Dim myC As SeqC: Set myC = SeqC(1, 2, 3, 4, 5, 6)
    myResult = Stringifier(e_WithTypes.m_Outer).ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Stringifier")
Private Sub Test05_StringifyItem_Collection()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{1,2,3,4,5,6}"
    
    Dim myResult As String
    Dim myC As Collection
    Set myC = New Collection
    myC.Add 1
    myC.Add 2
    myC.Add 3
    myC.Add 4
    myC.Add 5
    myC.Add 6
    'Act:  Again we need to sort The result SeqC to get the matching array

    myResult = Stringifier.ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Stringifier")
Private Sub Test05a_StringifyItemWithTypes_Collection()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Collection: {Integer: 1,Integer: 2,Integer: 3,Integer: 4,Integer: 5,Integer: 6}"
    
    Dim myResult As String
    Dim myC As Collection
    Set myC = New Collection
    myC.Add 1
    myC.Add 2
    myC.Add 3
    myC.Add 4
    myC.Add 5
    myC.Add 6
    'Act:  Again we need to sort The result SeqC to get the matching array

    myResult = Stringifier.Deb.WithTypes(e_WithTypes.m_All).ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test05b_StringifyItemWithTypesInner_Collection()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{Integer: 1,Integer: 2,Integer: 3,Integer: 4,Integer: 5,Integer: 6}"
    
    Dim myResult As String
    Dim myC As Collection
    Set myC = New Collection
    myC.Add 1
    myC.Add 2
    myC.Add 3
    myC.Add 4
    myC.Add 5
    myC.Add 6
    'Act:  Again we need to sort The result SeqC to get the matching array

    myResult = Stringifier(e_WithTypes.m_Inner).ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test05c_StringifyItemWithTypesOuter_Collection()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Collection: {1,2,3,4,5,6}"
    
    Dim myResult As String
    Dim myC As Collection
    Set myC = New Collection
    myC.Add 1
    myC.Add 2
    myC.Add 3
    myC.Add 4
    myC.Add 5
    myC.Add 6
    'Act:  Again we need to sort The result SeqC to get the matching array

    myResult = Stringifier(e_WithTypes.m_Outer).ResetMarkup.StringifyItem(myC)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test06_StringifyItem_Dictionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{ ""One"" 1, ""Two"" 2, ""Three"" 3, ""Four"" 4, ""Five"" 5, ""Six"" 6}"
    
    Dim myResult As String
    Dim myK As KvpC
    Set myK = KvpC.Deb
    myK.Add "One", 1
    myK.Add "Two", 2
    myK.Add "Three", 3
    myK.Add "Four", 4
    myK.Add "Five", 5
    myK.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    'Stringifier.ResetMarkup
    myResult = Stringifier.Deb.ResetMarkup.StringifyItem(myK)
    'Assert:
    'Debug.Print myExpected
    'Debug.Print myResult
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test06a_StringifyItemWithTypes_Dictionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "KvpC: { ""String: One"" Integer: 1, ""String: Two"" Integer: 2, ""String: Three"" Integer: 3, ""String: Four"" Integer: 4, ""String: Five"" Integer: 5, ""String: Six"" Integer: 6}"
    
    Dim myResult As String
    Dim myK As KvpC
    Set myK = KvpC.Deb
    myK.Add "One", 1
    myK.Add "Two", 2
    myK.Add "Three", 3
    myK.Add "Four", 4
    myK.Add "Five", 5
    myK.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    'Stringifier.ResetMarkup
    myResult = Stringifier.Deb.WithTypes(e_WithTypes.m_All).ResetMarkup.StringifyItem(myK)
    'Assert:
    ' Debug.Print
    ' Debug.Print myExpected
    ' Debug.Print myResult
  
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test06b_StringifyItemWithTypesInner_Dictionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{ ""String: One"" Integer: 1, ""String: Two"" Integer: 2, ""String: Three"" Integer: 3, ""String: Four"" Integer: 4, ""String: Five"" Integer: 5, ""String: Six"" Integer: 6}"
    
    Dim myResult As String
    Dim myK As KvpC
    Set myK = KvpC.Deb
    myK.Add "One", 1
    myK.Add "Two", 2
    myK.Add "Three", 3
    myK.Add "Four", 4
    myK.Add "Five", 5
    myK.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    'Stringifier.ResetMarkup
    myResult = Stringifier(e_WithTypes.m_Inner).ResetMarkup.StringifyItem(myK)
    'Assert:
    ' Debug.Print
    ' Debug.Print myExpected
    ' Debug.Print myResult
  
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test06c_StringifyItemWithTypesOuter_Dictionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "KvpC: { ""One"" 1, ""Two"" 2, ""Three"" 3, ""Four"" 4, ""Five"" 5, ""Six"" 6}"
    
    Dim myResult As String
    Dim myK As KvpC
    Set myK = KvpC.Deb
    myK.Add "One", 1
    myK.Add "Two", 2
    myK.Add "Three", 3
    myK.Add "Four", 4
    myK.Add "Five", 5
    myK.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    'Stringifier.ResetMarkup
    myResult = Stringifier(e_WithTypes.m_Outer).ResetMarkup.StringifyItem(myK)
    'Assert:
    ' Debug.Print
    ' Debug.Print myExpected
    ' Debug.Print myResult
  
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Stringifier")
Private Sub Test07_StringifyItem_CustomDictionaryMarkup()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{|One|_1?|Two|_2?|Three|_3?|Four|_4?|Five|_5?|Six|_6}"
    
    Dim myResult As String
    Dim myC As KvpC
    Set myC = KvpC.Deb
    myC.Add "One", 1
    myC.Add "Two", 2
    myC.Add "Three", 3
    myC.Add "Four", 4
    myC.Add "Five", 5
    myC.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    
    myResult = Stringifier.Deb.SetObjectMarkup(ipSeparator:="?").SetDictionaryItemMarkup("|", "_", "|").StringifyItembyKey(myC)
    'Assert:
    'Debug.Print myExpected
    'Debug.Print myResult
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test07a_StringifyItemWithTypes_CustomDictionaryMarkup()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "KvpC: {|String: One|_Integer: 1?|String: Two|_Integer: 2?|String: Three|_Integer: 3?|String: Four|_Integer: 4?|String: Five|_Integer: 5?|String: Six|_Integer: 6}"
    
    Dim myResult As String
    Dim myC As KvpC
    Set myC = KvpC.Deb
    myC.Add "One", 1
    myC.Add "Two", 2
    myC.Add "Three", 3
    myC.Add "Four", 4
    myC.Add "Five", 5
    myC.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    
    myResult = Stringifier.Deb.WithTypes(e_WithTypes.m_All).SetObjectMarkup(ipSeparator:="?").SetDictionaryItemMarkup("|", "_", "|").StringifyItembyKey(myC)
    'Assert:
    ' Debug.Print
    ' Debug.Print myExpected
    ' Debug.Print myResult
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test07b_StringifyItemWithTypesInner_CustomDictionaryMarkup()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{|String: One|_Integer: 1?|String: Two|_Integer: 2?|String: Three|_Integer: 3?|String: Four|_Integer: 4?|String: Five|_Integer: 5?|String: Six|_Integer: 6}"
    
    Dim myResult As String
    Dim myC As KvpC
    Set myC = KvpC.Deb
    myC.Add "One", 1
    myC.Add "Two", 2
    myC.Add "Three", 3
    myC.Add "Four", 4
    myC.Add "Five", 5
    myC.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    
    myResult = Stringifier(e_WithTypes.m_Inner).SetObjectMarkup(ipSeparator:="?").SetDictionaryItemMarkup("|", "_", "|").StringifyItembyKey(myC)
    'Assert:
    ' Debug.Print
    ' Debug.Print myExpected
    ' Debug.Print myResult
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test07c_StringifyItemWithTypesOuter_CustomDictionaryMarkup()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "KvpC: {|One|_1?|Two|_2?|Three|_3?|Four|_4?|Five|_5?|Six|_6}"
    
    Dim myResult As String
    Dim myC As KvpC
    Set myC = KvpC.Deb
    myC.Add "One", 1
    myC.Add "Two", 2
    myC.Add "Three", 3
    myC.Add "Four", 4
    myC.Add "Five", 5
    myC.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    
    myResult = Stringifier(e_WithTypes.m_Outer).SetObjectMarkup(ipSeparator:="?").SetDictionaryItemMarkup("|", "_", "|").StringifyItembyKey(myC)
    'Assert:
    ' Debug.Print
    ' Debug.Print myExpected
    ' Debug.Print myResult
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

