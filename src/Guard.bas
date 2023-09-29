Attribute VB_Name = "Guard"
'@Folder("Helpers")
Option Explicit

Public Enum e_GuardReportBackAction
    m_First = 0
    m_ReportBackContinue = e_GuardReportBackAction.m_First
    m_RaiseError
    m_Last = m_RaiseError
End Enum


Public Const MY_LIB                                     As String = "VBALib"
Public Const REPORT_BACK                                As Boolean = True

Private Type Properties
    ReportBackAction                                   As e_GuardReportBackAction
End Type

Private p                                               As Properties


Public Property Get ReportBackAction() As e_GuardReportBackAction
    ReportBackAction = p.ReportBackAction
End Property

Public Property Let ReportBackAction(ByVal ipReportBackAction As e_GuardReportBackAction)
    p.ReportBackAction = ipReportBackAction
End Property


Public Function IndexNotFound(ByVal ipIndex As Long, ByVal ipIndexed As Object, ByRef ipMethod As String, Optional ByVal ipReportBack As Boolean = False) As Boolean

    Dim myResult As Boolean: myResult = ipIndexed.LacksItems
    
    If Not myResult Then
        If ipIndex < ipIndexed.FirstIndex Or ipIndex > ipIndexed.LastIndex Then
            myResult = False
        End If
    End If
    
    If Not myResult Then
        IndexNotFound = myResult
        Exit Function
    End If
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipReportBack Then
            IndexNotFound = myResult
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipIndexed.TypeName, ipMethod), _
        Fmt.Text("Expecting a value between {0} and {1}. Got {2}.", ipIndexed.FirstIndex, ipIndexed.LastIndex, ipIndex)
    
End Function

' There are occasions when we wish to know the result of an alert test
' rather than trigger an error message
' In such cases ipREPORT_BACK should be set to true
Public Function IndexOutOfBounds(ByVal ipIndex As Long, ByRef ipKvp As Object, ByRef ipMethod As String, Optional ByRef ipReportBack As Boolean = False) As Boolean

    Dim myResult As Long: myResult = ((ipIndex < ipKvp.FirstIndex) Or (ipIndex > ipKvp.LastIndex))
    
    IndexOutOfBounds = myResult
    
    If Not myResult Then
        Exit Function
    End If
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipReportBack Then
            Exit Function
        End If
    End If

    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipKvp.TypeName, ipMethod), _
        Fmt.Text("Expecting a value between {0} and {1}. Got {2}.", ipKvp.FirstIndex, ipKvp.LastIndex, ipIndex)
    
End Function


Public Function InvalidRangeItem(ByRef ipRange As Variant, ByRef ipModule As String, ByRef ipMethod As String, Optional ByRef ipReportBack As Boolean = False) As Boolean

    Dim myResult As Boolean
    Select Case GroupInfo.Id(ipRange)
    
        Case e_Group.m_string, e_Group.m_array, e_Group.m_List, e_Group.m_Dictionary:       myResult = False
        Case Else:                                                                          myResult = True
            
    End Select
    
    InvalidRangeItem = myResult
    
    If Not myResult Then
        Exit Function
    End If
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipReportBack Then
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipModule, ipMethod), _
        Fmt.Text("Expecting a range item.  Got '{0}'", VBA.TypeName(ipRange))
    
End Function


Public Function ArrayNotFound(ByRef ipArray As Variant, ByRef ipModule As String, ByRef ipMethod As String, Optional ByVal ipReportBack As Boolean = False) As Boolean
    
    ArrayNotFound = ArrayOp.IsNotArray(ipArray)
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipReportBack Then
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipModule, ipMethod), _
        Fmt.Text("Expecting an array.  Got {0}", VBA.TypeName(ipArray))
        
End Function


Public Function EmptyRangeObject(ByRef ipRange As Variant, ByRef ipModule As String, ByRef ipMethod As String, Optional ByVal ipREPORT_BACK As Boolean = False) As Boolean

    Dim myLen As Long
    Select Case GroupInfo.Id(ipRange)
        Case e_Group.m_string:                  myLen = VBA.Len(ipRange)
        Case e_Group.m_array:                   myLen = ArrayOp.Count(ipRange)
        Case e_Group.m_List, m_Dictionary:      myLen = ipRange.Count
    End Select
    
    EmptyRangeObject = myLen < 1
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipREPORT_BACK Then
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipModule, ipMethod), _
        "Range object is empty"
        
End Function



Public Function KeyIsAdmin(ByRef ipAdmin As Variant, ByRef ipModule As String, ByRef ipMethod As String, Optional ByVal ipREPORT_BACK As Boolean = False) As Boolean

    Dim myResult As Boolean: myResult = GroupInfo.IsNotAdmin(ipAdmin)
    
    If myResult Then
        KeyIsAdmin = Not myResult
        Exit Function
    End If
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipREPORT_BACK Then
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipModule, ipMethod), _
        Fmt.Text("Admin values cannot be a Key. Got '{0}'", ipAdmin)
        
End Function

' todo: update As object to as IKvp when IKvp defined.
Public Function EnsureUniqueKeys(ByRef ipKey As Variant, ByVal ipKvp As Object, ByRef ipMethod As String, Optional ByVal ipReportBack As Boolean = False) As Boolean

    EnsureUniqueKeys = ipKvp.EnsureUniqueKeys
    
    If Not ipKvp.EnsureUniqueKeys Then
        Exit Function
    End If
    
    Dim myResult As Boolean: myResult = ipKvp.LacksKey(ipKey)
    
    If myResult Then
        EnsureUniqueKeys = myResult
        Exit Function
    End If
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipReportBack Then
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipKvp.TypeName, ipMethod), _
        Fmt.Text("Duplicate key.  Key '{0}' already exists", ipKey)
    
End Function

'@description("Guard to use when legitimately searching for a key as opposed to checking for the existance of a key")
Public Function KeyNotFound(ByVal ipLocOfKey As Variant, ByRef ipKey As Variant, ByRef ipComponent As String, ByRef ipMethod As String, Optional ByRef ipReportBack As Boolean = False) As Boolean
Attribute KeyNotFound.VB_Description = "Guard to use when legitimately searching for a key as opposed to checking for the existance of a key"
    
    If VBA.IsObject(ipLocOfKey) Then
        KeyNotFound = (ipLocOfKey Is Nothing)
    ElseIf GroupInfo.IsBoolean(ipLocOfKey) Then
        KeyNotFound = ipLocOfKey
    Else
        KeyNotFound = (ipLocOfKey = -1)
    End If
    
    If Not KeyNotFound Then
        Exit Function
    End If
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipReportBack Then
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipComponent, ipMethod), _
        Fmt.Text("Key not found. Got {0}:{1}", VBA.TypeName(ipKey), ipKey)

End Function

Public Function InvalidRun(ByRef ipRun As Long, ByRef ipComponent As String, ByRef ipMethod As String, Optional ByVal ipReportBack As Boolean = False) As Boolean

    InvalidRun = ipRun < 1
    
    If Not InvalidRun Then
        Exit Function
    End If
    
    If p.ReportBackAction = m_ReportBackContinue Then
        If ipReportBack Then
            Exit Function
        End If
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipComponent, ipMethod), _
        Fmt.Text("Runs of less than 1 are invalid.  Got {0}", ipRun)
        
End Function
