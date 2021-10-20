VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed


Private Type State
    
    Status                              As boolean
    Reason                              As ReasonCode
    Items                               As Lyst
    
End Type

Private s                               As State


Friend Function Deb() As Result
    
    With New Result
    
        Set Deb =.ReadyToUseInstance
        
    
    End With
    
    
End Function


Friend Function ReadyToUseInstance() As Result
    
    'Setup Result to be false with room for 5 items in the result
    ' Methods requiring multivalue results should use an 
    ' appropriate enumeration to access the result fields
    s.Status = False
    s.Reason = Failed
    s.Items = Lyst.Deb.AddValidatedIterable(Array(Empty, Empty, Empty, Empty, Empty))
    Set ReadyToUseInstance = Me
    
    
End Function

'@DefaultMember
Public Property Get Status() As Boolean
    Status = s.Status
End Property

Public Property Let Status(ByVal ipStatus As Boolean)
    s.Status = ipStatus
End Property



Public Property Get Reason() As Long
    Reason = s.Reason
End Property


Public Property Let Reason(ByVal ipReason As Long)
    s.Reason = ipReason
End Property


Public Property Get Item(ByVal ipIndex As Long) As Variant
    
    If VBA.IsObject(s.Items.Item(ipIndex)) Then
        
        Set Item = s.Items.Item(ipIndex)
        
        
    Else
        
       Item = s.Items.Item(ipIndex)
        
        
End If

End Property



Public Property Let Item(ByVal ipIndex As Long, ByVal ipItem As Variant)
    
    If VBA.IsObject(ipItem) Then
        
        Set s.Items.Item(ipIndex) = ipItem
        
    
    Else
        
        s.Items.Item(ipIndex) = ipItem
        
    End If
    
End Property


Public Sub Add(ByVal ipItem As Variant)
    s.Items.Add ipItem
End Sub

Public Function Failed() As Boolean
    Failed = Not s.Status
End Function


Public Function Okay() As Boolean
    Okay = s.Status
End Function

