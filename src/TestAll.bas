Attribute VB_Name = "TestAll"
'@IgnoreModule
'@Folder("Tests")
Option Explicit

Public myProcedureName As String
Public myComponentName As String

#If twinbasic Then
    ' Do Nothing
#Else

    ' We need these definitions because rubberduck unit testing is not being used.

    Public Assert As Object
    Public Fakes As Object

#End If

  Public Sub Main()
    #If twinbasic Then
        ' do nothing
    #Else
        ErrEx.Enable vbNullString
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #End If
        
    Dim myTime As Variant
    myTime = Timer
    Debug.Print "Testing started"
    Debug.Print
    TestArrayOp.ArrayOpTests
    TestFmt.FmtTests
    TestIterItems.IterItemsTests
    TestStrs.StrsTests
     TestStringifier.StringifierTests
     TestMappers.MapperTests
     TestComparers.ComparerTests
    'TestReducers.ReducerTests
    TestIterItems.IterItemsTests

'    TestHashC.cHashCTests  ' not yet complete
    TestSeqA.SeqATests
    TestSeqC.SeqCTests
    TestSeqL.SeqLTests

    TestKvpA.KvpATests
    TestKvpC.KvpCTests
    TestKvpL.KvpLTests
    'TestKvpH.KvpHTests
    Debug.Print
    Debug.Print "Testing Finished  " & Timer - myTime & " seconds."
        
End Sub

'Private Function TimeNow() As String
'    TimeNow = Format(Time(), "hh:nn:ss")
'End Function

