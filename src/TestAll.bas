Attribute VB_Name = "TestAll"
'@IgnoreModule
'@Folder("Tests")
Option Explicit

Public myProcedureName As String
Public myComponentName As String
Public VBATesting   As Boolean
#If twinbasic Then
    ' Do Nothing
#Else
    Public Assert As Object
    Public Fakes As Object
#End If


Public Sub Main()
  
    VBATesting = True
    Dim myTime As Variant: myTime = Timer
     
    Debug.Print "Testing started"
    Debug.Print

    TestArrayOp.ArrayOpTests
    TestFmt.FmtTests
    TestIterItems.IterItemsTests
    TestStrs.StrsTests
    TestStringifier.StringifierTests
    TestMappers.MapperTests
    TestCmpFunctors.CmpFunctorTests
    'TestReducers.ReducerTests
    '    TestHashC.cHashCTests  ' not yet complete
    TestSeqA.SeqATests
    TestSeqC.SeqCTests
    TestSeqL.SeqLTests
    TestSeqH.SeqHTests

    TestKvpA.KvpATests
    TestKvpC.KvpCTests
    TestKvpH.KvpHTests
    TestKvpL.KvpLTests
    VBATesting = False
    
    Debug.Print
    Debug.Print "Testing Finished  " & Timer - myTime & " seconds."
        
End Sub

Public Sub GlobalAssert()
    #If twinbasic Then
        ' do nothing
    #Else
        If Not ErrEx.IsEnabled Then
            ErrEx.Enable vbNullString
        End If
        If Assert Is Nothing Then
            Set Assert = New Rubberduck.AssertClass
        End If
        If Fakes Is Nothing Then
            Set Fakes = New Rubberduck.FakesProvider
        End If
    #End If

End Sub
