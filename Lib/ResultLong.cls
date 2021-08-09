VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultLong"
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
    value                           As Long
    
End Type

Private p                           As Properties

'@Description("Returns a new instance prepoulated with fail values")
Public Function Deb _
( _
    Optional ByVal ipStatus As Boolean = False, _
    Optional ByVal ipReason As ReasonCode = rcfail, _
    Optional ByVal ipWhy As String = vbNullString, _
    Optional ByVal ipValue As Long _
) As ResultLong
Attribute Deb.VB_Description = "Returns a new instance prepoulated with fail values"

    If Me Is ResultLong Then
    
        With New ResultLong
        
            Set Deb = .ReadyToUseInstance(ipStatus, ipReason, ipWhy, ipValue)
        
        End With
    
    End If
    
End Function

Public Function ReadyToUseInstance _
( _
    ByVal ipStatus As Boolean, _
    ByVal ipReason As ReasonCode, _
    ByVal ipWhy As String, _
    Optional ByVal ipValue As Variant _
) As ResultLong

    If Me Is ResultLong Then
    
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
        .value = IIf(VBA.IsMissing(ipValue), -1&, CLng(ipValue))
        
    End With
    
    Set ReadyToUseInstance = Me
    
End Function
    
    
' All properties are read/write
' This is because
'  set opresult = ResultLong.deb
'  Set tryget<thing> = opResult
' assigns to tryGet<thing> the reference of the object created
' by resultLong.deb and not a reference to opResult.
'
' This if we then do
' set opresult = resultlong.SetSuccess(value)
' then opresult and TryGet<thing> refer to different objects
' TryGet<thing> will not refer to the object created by resultlong.SetSuccess
' but will keep the reference to resultlong.deb

'@DefaultMember
Public Property Get Status() As Boolean
Attribute Status.VB_UserMemId = 0
    Status = p.Status
End Property

Public Property Let Status(ByVal ipStatus As Boolean)
    p.Status = ipStatus
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

Public Property Let value(ByVal ipValue As Variant)
    p.value = CLng(ipValue)
End Property

Public Function Failed() As Boolean
    Failed = Not p.Status
End Function

Public Function Succeeded() As Boolean
    Succeeded = p.Status
End Function

Public Function SetFailure(Optional ByVal ipReason As ReasonCode = rcfail, Optional ByVal ipString As String = vbNullString, Optional ByVal ipValue As Long = -1) As ResultLong

    p.Status = False
    p.Reason = ipReason
    p.value = ipValue
    
    If VBA.Len(ipString) > 0 Then
        
        p.Why = ipString
        
    Else
        
        p.Why = ResultReason.AsString(ipReason)
        
    End If
    
    Set SetFailure = Me
        
End Function

Public Function SetSuccess(Optional ByVal ipValue As Long = -1) As ResultLong
    
    p.Status = True
    p.Reason = rcSuccess
    p.Why = ResultReason.AsString(rcSuccess)
    p.value = ipValue
    
    Set SetSuccess = Me
        
End Function

Public Sub Copy(ByVal ipResult As ResultLong)

    p.Status = ipResult.Status
    p.Reason = ipResult.Reason
    p.Why = ipResult.Why
    p.value = ipResult.value

End Sub
