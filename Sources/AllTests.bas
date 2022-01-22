Attribute VB_Name = "AllTests"
Option Explicit
'@IgnoreModule
'@TestModule
'@Folder("Tests")


Public myStart As Variant
Public myEnd As Variant
Public myInterim As Variant

Public Sub Tester()

    ErrEx.Enable ""
    myStart = Timer
    Debug.Print "Testing started"
    
    ' Do test in this order or suffer pain
    TestStringifier.StringifierTests
    TestFmt.FmtTests
    TestTYpes.TypesTests
    TestStrs.StrsTests
    TestRanges.RangeTests
    TestArrays.ArraysTests
    TestLyst.LystTests
    TestKvp.KvpTests

    myEnd = Timer
    Debug.Print "Testing completed in " & myEnd - myStart & " seconds"
    
End Sub

