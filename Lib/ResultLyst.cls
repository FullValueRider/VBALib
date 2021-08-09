VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultLyst"
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
    value                           As Lyst
    
End Type

Private p                           As Properties

Public Function Deb _
( _
    Optional ByVal ipStatus As Boolean = False, _
    Optional ByVal ipReason As ReasonCode = rcfail, _
    Optional ByVal ipWhy As String = vbNullString, _
    Optional ByVal ipValue As Lyst = Nothing _
) As ResultLyst

    If Me Is ResultLyst Then
    
        With New ResultLyst
        
            Set Deb = .ReadyToUseInstance(ipStatus, ipReason, ipWhy, ipValue)
        
        End With
        
    End If
    
End Function

Public Function ReadyToUseInstance _
( _
    ByVal ipStatus As Boolean, _
    ByVal ipReason As ReasonCode, _
    ByVal ipWhy As String, _
    ByVal ipValue As Lyst _
) As ResultLyst

    If Me Is ResultLyst Then
    
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
        
        If ipValue Is Nothing Then
        
            Set .value = Lyst.Deb
            
        Else
        
            Set .value = ipValue
            
        End If
        
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


Public Property Get value() As Lyst
    Set value = p.value
End Property

Public Property Set value(ByVal ipValue As Lyst)
    Set p.value = ipValue
End Property

'Sugar for status which gives true if operation fails
Public Function Failed() As Boolean
    Failed = Not p.Status
End Function

'SUgar for Status which give true is operation succeeds
Public Function Succeeded() As Boolean
    Succeeded = p.Status
End Function

Public Function SetFailure(Optional ByVal ipReason As ReasonCode = rcfail, Optional ByVal ipWhy As String = vbNullString, Optional ByVal ipValue As Lyst = Nothing) As ResultLyst

    p.Status = False
    p.Reason = ipReason
    p.Why = ipWhy
    Set p.value = ipValue
    Set SetFailure = Me
        
End Function

Public Function SetSuccess(ByVal ipValue As Lyst) As ResultLyst
    
    p.Status = True
    p.Reason = rcSuccess
    p.Why = vbNullString
    Set p.value = ipValue
    
    Set SetSuccess = Me
        
End Function

Public Sub Copy(ByVal ipResult As ResultLyst)

    p.Status = ipResult.Status
    p.Reason = ipResult.Reason
    p.Why = ipResult.Why
    Set p.value = ipResult.value

End Sub
