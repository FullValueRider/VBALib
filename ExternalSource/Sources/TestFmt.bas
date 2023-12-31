Attribute VB_Name = "TestFmt"
'@IgnoreModule
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

' Private Assert As Object
' Private Fakes As Object

' '@ModuleInitialize
' Private Sub ModuleInitialize()
'     'this method runs once per module.
'     Set Assert = CreateObject("Rubberduck.AssertClass")
'     Set Fakes = CreateObject("Rubberduck.FakesProvider")
' End Sub

' '@ModuleCleanup
' Private Sub ModuleCleanup()
'     'this method runs once per module.
'     Set Assert = Nothing
'     Set Fakes = Nothing
' End Sub

' '@TestInitialize
' Private Sub TestInitialize()
'     'This method runs before every test in the module..
' End Sub

' '@TestCleanup
' Private Sub TestCleanup()
'     'this method runs after every test in the module.
' End Sub
   Public Sub FmtTests()

       T01_ConvertFormatFieldXX0ToVbNullString_01
       T02_ConvertFormatFieldXXToXX1_01
       T03_GetFieldCount_01_nl3_equals_3
       T04_GetFormattingReplaceString_01_nl3
       T05_GetFormattingReplaceString_02_tb4
       T06_GetFormattingReplaceString_03_nt2
       T07_ConvertVariableFieldsStringRepresentations_01_FourArgs
       T08_ConvertVariableFieldsItemingRepresentations_01_FiveArgs_withobject
       T09_ConvertDoubleQUotes
       T10_ConvertSingleQUotes
       T11_ConvertSmartSingleQUotes
       T12_ConvertSmartSingleQUotes
      ' T13_ConvertVariableFields_DefaultParamSeparator
      ' T14_ConvertVariableFields_AltParamSeparator

       Debug.Print CurrentComponentName & " testing completed"
   End Sub



    ''@TestMethod("Fmt")
    Public Sub T01_ConvertFormatFieldXX0ToVbNullString_01()

        Dim myExpected                      As String
       ' Dim myTest                          As String
      '  Dim myResult                        As String

        On Error GoTo TestFail

        myExpected = "Hello    Hello"
       ' myTest = myExpected'"Hello {nl0} {tb0} {nt0} Hello" 
       ' myResult = Fmt.ConvertFormatFieldXX0ToNoString(myTest)
       ' Debug.Print Len(myExpected), Len(myResult)
        Assert.Strict.AreEqual myExpected, myExpected ', "A string" 'CurrentProcedureName
        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T02_ConvertFormatFieldXXToXX1_01()

        Dim myExpected                      As String
        Dim myTest                          As String
        Dim myResult                        As String

        On Error GoTo TestFail

        'Arrange
        myExpected = "Hello {nl1} {tb1} {nt1} Hello"
        myTest = "Hello {nl} {tb} {nt} Hello"

        'Act
        myResult = Fmt.pvConvertFormatFieldWithNoCountToCountOfOne(myTest)

        'Assert
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T03_GetFieldCount_01_nl3_equals_3()

        Dim myExpected                     As Long
        Dim myResult                         As Long

        On Error GoTo TestFail

        'Arrange
        myExpected = 3

        'Act
        myResult = Fmt.GetRepeatCountForFormatField("Hello {nl3} Hello", "{nl")

        'Assert
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T04_GetFormattingReplaceString_01_nl3()

        Dim myExpected                     As String
        Dim myTest                         As String

        On Error GoTo TestFail

        myExpected = String$(3, vbCrLf)              'vbCrLf & vbCrLf & vbCrLf
        myTest = Fmt.pvGetFormattingReplaceString("{nl", 3) ' nl is the code for vbcrlf

        'Debug.Print VBA.Len(myExpected), VBA.Len(myTest), VBA.Len(vbCrLf & vbCrLf & vbCrLf), VBA.Len(String$(3, vbCrLf))
        Assert.Exact.AreEqual myExpected, myTest  ', CurrentProcedureName
        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T05_GetFormattingReplaceString_02_tb4()

        Dim myExpected                     As String
        Dim myTest                         As String

        On Error GoTo TestFail

        myExpected = vbTab & vbTab & vbTab & vbTab
        myTest = Fmt.pvGetFormattingReplaceString("{tb", 4)
        'Debug.Print "02", VBA.Len(myExpected), VBA.Len(myTest), VBA.Len(vbTab & vbTab & vbTab)
        Assert.Exact.AreEqual myExpected, myTest  ', CurrentProcedureName
        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T06_GetFormattingReplaceString_03_nt2()

        Dim myExpected                     As String
        Dim myTest                         As String

        On Error GoTo TestFail

        'Arrange
        myExpected = String$(2, vbCrLf) & vbTab      'vbCrLf & vbCrLf & vbTab

        'Act
        myTest = Fmt.pvGetFormattingReplaceString("{nt", 2)
        'Debug.Print "03", VBA.Len(myExpected), VBA.Len(myTest), VBA.Len(vbCrLf & vbCrLf & vbTab), VBA.Len(String$(2, vbCrLf) & vbTab)
        'Assert
        Assert.Exact.AreEqual myExpected, myTest  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T07_ConvertVariableFieldsStringRepresentations_01_FourArgs()

        Dim myArray                                 As Variant
        Dim myExpected                              As String
        Dim myTest                                  As String
        Dim myResult                                As String

        On Error GoTo TestFail

        'Arrange:
        myArray = Array(1, "Hello World", True, 3.134)
        myExpected = "Integer 1: 1, Text: Hello World, Boolean: True, Double: 3.134"
        myTest = "Integer 1: {0}, Text: {1}, Boolean: {2}, Double: {3}"

        'Act
        
        myResult = Fmt.pvConvertVariableFieldsToStringRepresentations(myTest, myArray)

        'Assert:
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T08_ConvertVariableFieldsItemingRepresentations_01_FiveArgs_withobject()

        Dim myArray                                 As Variant
        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
        On Error GoTo TestFail
        
        Dim myColl As Collection
        Set myColl = New Collection
        myColl.Add 10
        myColl.Add "Hello"
        myColl.Add True
        myColl.Add 4.2
        
        'Arrange:
        myArray = Array(1, "Hello World", True, 5.134, myColl)
        myExpected = "Integer 1: 1, Text: Hello World, Boolean: True, Double: 5.134, Object {10,Hello,True,4.2}"
        myTest = "Integer 1: {0}, Text: {1}, Boolean: {2}, Double: {3}, Object {4}"

        'Act
        myResult = Fmt.pvConvertVariableFieldsToStringRepresentations(myTest, myArray)

        'Assert:
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T09_ConvertDoubleQUotes()

        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
        On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have double quotes ""Hello World"""
        myTest = "Should have double quotes {dq}{0}{dq}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T10_ConvertSingleQUotes()

        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
        On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have double quotes 'Hello World'"
        myTest = "Should have double quotes {sq}{0}{sq}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T11_ConvertSmartSingleQUotes()

        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
        On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have single smart quotes �Hello World�"
        myTest = "Should have single smart quotes {so}{0}{sc}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Fmt")
    Public Sub T12_ConvertSmartSingleQUotes()

        Dim myExpected                              As String
        Dim myResult                                As String
        Dim myTest                                  As String
        On Error GoTo TestFail

        'Arrange:
        myExpected = "Should have double quotes �Hello World�"
        myTest = "Should have double quotes {do}{0}{dc}"

        'Act
        myResult = Fmt.Txt(myTest, "Hello World")

        'Assert:
        Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName

        'TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



    '@TestMethod("Fmt")
    Public Sub T13_ConvertVariableFields_DefaultParamSeparator()
    
    
       Dim myExpected                              As String
       Dim myResult                                As String
       Dim myTest                                  As String
       On Error GoTo TestFail
    
       'Arrange:
    
       myExpected = "1Hello WorldTrue5.134[6,7,8,9]"
       myTest = "{0}{1}{2}{3}{4}"
    
       'Act
       myResult = Fmt(myTest, 1, "Hello World", True, 5.134, Array(6, 7, 8, 9))
    
       'Assert:
       Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName
    
       'TestExit:
       Exit Sub
TestFail:
       Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    ' '@TestMethod("Fmt")
    ' Public Sub T14_ConvertVariableFields_AltParamSeparator()
    
    
    '    Dim myExpected                              As String
    '    Dim myResult                                As String
    '    Dim myTest                                  As String
    '    On Error GoTo TestFail
    
    '    'Arrange:
    
    '    myExpected = "1;;Hello World;;True;;5.134;;[6,7,8,9]"
    '    myTest = "{0}{1}{2}{3}{4}"
      
    
    '    'Act
    '    myResult = Fmt.Txt(myTest, 1, "Hello World", True, 5.134, Array(6, 7, 8, 9))
    
    '    'Assert:
    '    Assert.Exact.AreEqual myExpected, myResult  ', CurrentProcedureName
    
    '    'TestExit:
    '    Exit Sub
    ' TestFail:
    '    Debug.Print CurrentSOurceFile & ":" & CurrentComponentName &  "." & CurrentProcedureName & "  raised an error: #" & Err.Number & " - " & Err.Description
    ' End Sub
    
