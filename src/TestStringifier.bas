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
        Debug.Print CurrentProcedureName; vbTab, vbTab, vbTab,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    #End If

    Test01_StringifyItem_String
    Test02_StringifyItem_Long
    Test03_StringifyItem_Array
    Test04_StringifyItem_SeqC
    Test05_StringifyItem_Collection
    Test06_StringifyItem_Dictionary
    Test07_StringifyItem_CustomDictionaryMarkup
    
    
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
    myResult = Stringifier.StringifyItem("Hello World!")
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
    myResult = Stringifier.StringifyItem(42)
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


''@TestMethod("Stringifier.049")
'Private Sub Test49c_StringifyItem_Long()
'
'    on error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As String
'    myExpected = "42"
'
'    Dim myresult As String
'
'    'Act:  Again we need to sort The result SeqC to get the matching array
'    myresult = Stringifier.StringifyItem(42)
'    'Assert:
'    AssertExactAreEqual myExpected, myResult,myProcedureName
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    on error Resume Next
'
'    Exit Sub
'TestFail:
'    AssertFail myCOmponentName, myProcedureName," raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub


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
    myResult = Stringifier.StringifyItem(Array(1, 2, 3, 4, 5, 6))
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
    myResult = Stringifier.StringifyItem(myC)
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
    Stringifier.ResetMarkup
    myResult = Stringifier.StringifyItem(myC)
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
    myExpected = "{ 'One': 1, 'Two': 2, 'Three': 3, 'Four': 4, 'Five': 5, 'Six': 6}"
    
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
    myResult = Stringifier.StringifyItem(myK)
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
    Dim myToString As Stringifier
    Set myToString = Stringifier.Deb.SetObjectMarkup(ipSeparator:="?").SetDictionaryItemMarkup("|", "_", "|")
    myResult = myToString.StringifyItem(myC)
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


