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
    TestFormat.FormatTests
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
    TestSeqHC.SeqHCTests
    TestSeqHL.SeqHLTests
    'TestSeqT.SeqTTests
    TestKvpA.KvpATests
    TestKvpC.KvpCTests
    TestKvpHA.KvpHATests
    tESTkVPhl.KvpHLTests
    TestKvpL.KvpLTests
    TestKvpLP.KvpLPTests
    
    'AoC CLasses
    TestPointXY.PointXYTests
    'TestTrackedXY.TrackedXYTests
    
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
