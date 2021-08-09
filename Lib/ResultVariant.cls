VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")

Private Type Properties

    Status                          As Boolean
    Reason                          As ReasonCode
    Why                             As String
    value                           As Variant
    
End Type

Private p                           As Properties

Public Function Deb _
( _
    Optional ByVal ipStatus As Boolean = False, _
    Optional ByVal ipReason As ReasonCode = ReasonCode.rcfail, _
    Optional ByVal ipWhy As String = vbNullString, _
    Optional ByVal ipValue As Variant = Empty _
) As ResultVariant

    If Me Is ResultVariant Then
    
        With New ResultVariant
        
            Set Deb = .ReadyToUseInstance(ipStatus, ipReason, ipWhy, ipValue)
        
        End With
    
    End If
    
End Function

Public Function ReadyToUseInstance _
( _
    ByVal ipStatus As Boolean, _
    ByVal ipReason As ReasonCode, _
    ByVal ipWhy As String, _
    ByVal ipValue As Variant _
) As ResultVariant

    If Me Is ResultVariant Then
    
        Err.Raise _
            vbObjectError + 17, _
            "VBALib.ResultLong.ReadyToUseInstance", _
            "ReadyToUseInstance is not allowed for the PredeclaredId"
            
        Exit Function
    
    End If
    
    With p
    
        .Status = ipStatus
        .Reason = ipReason
        .Why = ipWhy
        .value = ipValue
        
    End With
    
    Set ReadyToUseInstance = Me
    
End Function
    
'@DefaultMember
Public Property Get Status() As Boolean
Attribute Status.VB_UserMemId = 0
    Status = p.Status
End Property

Public Property Get Reason() As ReasonCode
    Reason = p.Reason
End Property

Public Property Let Reason(ByVal ipReason As ReasonCode)
    p.Reason = ipReason
End Property


Public Property Get Why() As String
    Why = p.Why
End Property

Public Property Let Why(ByVal ipWhy As String)
    p.Why = ipWhy
End Property

Public Property Get value() As Variant
    value = p.value
End Property

Public Property Set value(ByVal ipValue As Variant)

    p.value = ipValue
    
End Property


Public Function Failed() As Variant
    Failed = Not p.Status
End Function

Public Function Succeeded() As Boolean
    Succeeded = p.Status
End Function

Public Function SetFailure(Optional ByVal ipReason As ReasonCode = rcfail, Optional ByVal ipWhy As String = vbNullString, Optional ByVal ipValue As Variant = Empty) As ResultVariant

    p.Status = False
    p.Reason = ipReason
    p.Why = ipWhy
    p.value = ipValue
    Set SetFailure = Me
        
End Function

Public Function SetSuccess(ByVal ipValue As Variant) As ResultVariant
    
    p.Status = True
    p.Reason = rcSuccess
    p.Why = vbNullString
    p.value = ipValue
    
    Set SetSuccess = Me
        
End Function

Public Sub Copy(ByVal ipResult As ResultVariant)

    p.Status = ipResult.Status
    p.Reason = ipResult.Reason
    p.Why = ipResult.Why
    p.value = ipResult.value

End Sub

