Attribute VB_Name = "Guard"
'@Folder("Helpers")
Option Explicit

Public Const MY_LIB                                     As String = "VBALib"
Public Const REPORT_BACK                                As Boolean = True
Public Enum e_Abort

    m_First = -1
    m_KeyIsMissing = m_First
    m_InvalidIndex = m_First
    m_Last
End Enum


' There are occasions when we wish to know the result of an alert test
' rather than trigger an error message
' In such cases ipREPORT_BACK should be set to true
Public Function IndexOutOfBounds(ByVal ipIndex As Long, ByRef ipKvp As Object, ByRef ipMethod As String, Optional ByRef ipReportBack As Boolean = False) As Boolean

    Dim myResult As Long: myResult = ((ipIndex < ipKvp.FirstIndex) Or (ipIndex > ipKvp.LastIndex))
    
    IndexOutOfBounds = myResult
    
    If Not myResult Then
        Exit Function
    End If
    
    If ipReportBack Then
    
        Exit Function
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
    
    If ipReportBack Then
        Exit Function
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipModule, ipMethod), _
        Fmt.Text("Expecting a range item.  Got '{0}'", VBA.TypeName(ipRange))
    
End Function


'Public Sub GuardInvalidIndex(ByRef ipIndex As Long, ByRef ipLastIndex As Long, ByRef ipMessage As String)
'
'    If ipIndex <= ipLastIndex Then
'        Exit Sub
'    End If
'
'    Err.Raise 17 + vbObjectError, _
'    ipMessage, _
'    Fmt.Text("Index {0} is not available.  Capacity is {1}", ipIndex, ipLastIndex)
'End Sub


'Public Sub GuardInsufficientCapacity(ByRef ipInitialSize As Long, ByRef ipMessage As String)
'
'    If ipInitialSize < 1 Then
'        Err.Raise 17 + vbObjectError, _
'        ipMessage, _
'        Fmt.Text("Got initial size of {0} . Expecting a positive integer greater than 0", VBA.CStr(ipInitialSize))
'    End If
'
'End Sub


Public Function ArrayNotFound(ByRef ipArray As Variant, ByRef ipModule As String, ByRef ipMethod As String, Optional ByVal ipREPORT_BACK As Boolean = False) As Boolean
    
    ArrayNotFound = ArrayOp.IsNotArray(ipArray)
    
    If ipREPORT_BACK Then
        Exit Function
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
    
    If ipREPORT_BACK Then
        Exit Function
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
    
    
    If ipREPORT_BACK Then
        Exit Function
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipModule, ipMethod), _
        Fmt.Text("Admin values cannot be a Key. Got '{0}'", ipAdmin)
        
End Function

' todo: update As object to as IKvp when IKvp defined.
Public Function EnsureUniqueKeys(ByRef ipKvp As Object, ByRef ipKey As Variant, ByRef ipModule As String, ByRef ipMethod As String, Optional ByVal ipReportBack As Boolean = False) As Boolean

    EnsureUniqueKeys = ipKvp.EnsureUniqueKeys
    
    If Not ipKvp.EnsureUniqueKeys Then
        Exit Function
    End If
    
    Dim myResult As Boolean: myResult = ipKvp.LacksKey(ipKey)
    
    If myResult Then
        EnsureUniqueKeys = myResult
        Exit Function
    End If
    
    If ipReportBack Then
        Exit Function
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipModule, ipMethod), _
        Fmt.Text("Duplicate key.  Got '{0}'", ipKey)
    
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
    
    If ipReportBack Then
        Exit Function
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
    
    If ipReportBack Then
        Exit Function
    End If
    
    Err.Raise 17 + vbObjectError, _
        Fmt.Text("{0}.{1}.{2}", MY_LIB, ipComponent, ipMethod), _
        Fmt.Text("Runs of less than 1 are invalid.  Got {0}", ipRun)
        
End Function
