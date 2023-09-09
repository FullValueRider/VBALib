Attribute VB_Name = "TestSeqPerformance"
Option Explicit

' add your procedures here
'Private Declare PtrSafe Function QueryPerformanceFrequency& Lib "kernel32" (x@)
'Private Declare PtrSafe Function QueryPerformanceCounter& Lib "kernel32" (x@)
'
Private Const TestEntryCount As Long = 20000
    
     
Public Sub Main()
Dim i As Long
Dim T As Single
'@Ignore VariableNotUsed
Dim Item As Variant
  
  Debug.Print " Count of Test-Entries:"; TestEntryCount; vbLf
    
    T = VBA.Timer
    Dim mySA As SeqA: Set mySA = SeqA.Deb
    Debug.Print "Initialising SeqA", VBA.Timer - T & "msec"
    T = VBA.Timer
    Dim mySC As SeqC: Set mySC = SeqC.Deb
    Debug.Print "Initialising SeqC", VBA.Timer - T & "msec"
    T = VBA.Timer
    Dim mySL As SeqL: Set mySL = SeqL.Deb
    Debug.Print "Initialising SeqL", VBA.Timer - T & "msec"
    T = VBA.Timer
    Dim mySH As SeqHL: Set mySH = SeqHL.Deb
    mySH.Reinit 32767
    Debug.Print "Initialising SeqL", VBA.Timer - T & "msec"
    T = VBA.Timer
    Dim myTC As Collection: Set myTC = New Collection
    Debug.Print "Initialising COllection", VBA.Timer - T & "msec"
    

    Debug.Print
    '@Ignore AssignmentNotUsed
    T = VBA.Timer
    
    
    Debug.Print
    Debug.Print "Adding String Items"
    Debug.Print
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      mySA.Add VBA.CStr(i)
    Next
  Debug.Print "SeqA Add:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      mySC.Add VBA.CStr(i)
    Next
  Debug.Print "SeqC Add:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      mySL.Add VBA.CStr(i)
    Next
  Debug.Print "SeqL Add:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      mySH.Add VBA.CStr(i)
    Next
  Debug.Print "SeqH Add:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      myTC.Add VBA.CStr(i)
    Next
  Debug.Print "Collection Add:", VBA.Timer - T & "msec"
  
  Debug.Print
  Debug.Print
  Debug.Print
  
  Debug.Print "Get Item"
  Debug.Print
  Debug.Print
  Debug.Print
  
  
 
  T = VBA.Timer
    For i = 1 To TestEntryCount
      Item = mySA.Item(i)
    Next
  Debug.Print "SeqA Item:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      Item = mySC.Item(i)
    Next
  Debug.Print "SeqC Item:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      Item = mySL.Item(i)
    Next
  Debug.Print "SeqL Item:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      Item = mySH.Item(i)
    Next
  Debug.Print "SeqH Item:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      Item = myTC.Item(i)
    Next
  Debug.Print "Collection Item:", VBA.Timer - T & "msec"
  
  
  Debug.Print
  Debug.Print
  Debug.Print
  Debug.Print " HoldsItem"
  Debug.Print
  Debug.Print
  Debug.Print
  
  '@Ignore VariableNotUsed
  Dim bHoldsItem As Boolean
  
'  T = VBA.Timer
'    For i = 1 To TestEntryCount
'      bHoldsItem = mySA.HoldsItem(VBA.CStr(i))
'    Next
'  Debug.Print "SeqA HoldsItem:", VBA.Timer - T & "msec"
'
'  T = VBA.Timer
'    For i = 1 To TestEntryCount
'      bHoldsItem = mySC.HoldsItem(VBA.CStr(i))
'    Next
'  Debug.Print "SeqC HoldsItem:", VBA.Timer - T & "msec"
  
  T = VBA.Timer
    For i = 1 To TestEntryCount
      bHoldsItem = mySL.HoldsItem(VBA.CStr(i))
    Next
  Debug.Print "SeqL HoldsItem:", VBA.Timer - T & "msec"
      
  T = VBA.Timer
    For i = 1 To TestEntryCount
      bHoldsItem = mySH.HoldsItem(VBA.CStr(i))
    Next
  Debug.Print "SeqH HoldsItem:", VBA.Timer - T & "msec"
  
'   T = VBA.Timer
'     For i = 1 To TestEntryCount
'       bHoldsItem = myTC.Exists(VBA.CStr(i))
'     Next
'   Debug.Print "tb collection HoldsItem:", VBA.Timer - T & "msec"
  
mySH.Analyse
  
  'H.CheckHashDistribution
End Sub
    
'Function VBA.Timer@()
'Dim x@, Frq@
'  QueryPerformanceFrequency Frq
'  If QueryPerformanceCounter(x) Then VBA.Timer = CCur(x / Frq) * 1000
'End Function
