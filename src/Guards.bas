Attribute VB_Name = "Guards"
'@Folder("Helpers")
Option Explicit


Public Sub GuardIndexOutOfBounds(ByRef ipIndex As Long, ByRef ipLowerIndex As Long, ByRef ipUpperIndex As Long, ByRef ipMessage As String)

    If ipIndex >= ipLowerIndex And ipIndex <= ipUpperIndex Then
        Exit Sub
    End If

    Err.Raise 17 + vbObjectError, _
    ipMessage, _
    Fmt.Text("Expecting a value between {0} and {1}. Got {2}.", ipLowerIndex, ipUpperIndex, ipIndex)
    
End Sub


Public Sub GuardInvalidRangeItem(ByRef ipRange As Variant, ByRef ipMessage As String)

    Select Case GroupInfo.Id(ipRange)
    
        Case e_Group.m_string:      Exit Sub
        Case e_Group.m_array:       Exit Sub
        Case e_Group.m_List:        Exit Sub
        Case e_Group.m_Dictionary:  Exit Sub
        Case Else
            Err.Raise 17 + vbObjectError, _
            ipMessage, _
            Fmt.Text("Expecting string, array, list type or dictionary type.  Got {0}", VBA.TypeName(ipRange))
    End Select
    
End Sub


Public Sub GuardInvalidIndex(ByRef ipIndex As Long, ByRef ipLastIndex As Long, ByRef ipMessage As String)

    If ipIndex <= ipLastIndex Then
        Exit Sub
    End If
    
    Err.Raise 17 + vbObjectError, _
    ipMessage, _
    Fmt.Text("Index {0} is not available.  Capacity is {1}", ipIndex, ipLastIndex)
End Sub


Public Sub GuardInsufficientCapacity(ByRef ipInitialSize As Long, ByRef ipMessage As String)

    If ipInitialSize < 1 Then
        Err.Raise 17 + vbObjectError, _
        ipMessage, _
        Fmt.Text("Got initial size of {0} . Expecting a positive integer greater than 0", VBA.CStr(ipInitialSize))
    End If
    
End Sub


Public Sub GuardExpectingArray(ByRef ipArray As Variant, ByRef ipMessage As String)
    
    If ArrayOp.IsNotArray(ipArray) Then
        Err.Raise 17 + vbObjectError, _
        ipMessage, _
        Fmt.Text("Expecting an array.  Got {0}", VBA.TypeName(ipArray))
    End If
        
End Sub


Public Sub GuardEmptyRangeObject(ByRef ipRange As Variant, ByRef ipMessage As String)

    Dim myLen As Long
    Select Case GroupInfo.Id(ipRange)
        Case e_Group.m_string:                  myLen = VBA.Len(ipRange)
        Case e_Group.m_array:                   myLen = ArrayOp.Count(ipRange)
        Case e_Group.m_List, m_Dictionary:      myLen = ipRange.Count
    End Select
    
    If myLen < 1 Then
        Exit Sub
    End If
    
    Err.Raise 17 + vbObjectError, _
    ipMessage, _
    "Range object is empty"
        
End Sub



