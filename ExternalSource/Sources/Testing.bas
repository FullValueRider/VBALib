Attribute VB_Name = "Testing"
 '@IgnoreModule
Option Explicit

Public myPlace As string
'@TestModule
'@Folder("Tests")

' Private Assert As Object
' Private Fakes As Object

' '@ModuleInitialize
' Private Sub ModuleInitialize()
'     'this method runs once per module.
'     myPlace = "VBALib"
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
    'Public Assert As AssertClass
    Public Sub Tester()
        'Dim assert As AssertClass
       ' Set Assert = New AssertClass
        
        Debug.Print "Testing started"
         
       'TestStringifier.StringifierTests
       'TestFmt.FmtTests
        TestLyst.LystTests
    '    TestArrays.ArraysTests
    '    TestTypes.TypesTests
        'TestKvp.KvpTests
        
        Debug.Print "Testing completed"
        
    End Sub
