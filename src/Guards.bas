Attribute VB_Name = "Guards"
Option Explicit

Public Sub GuardIndexOutOfBounds(ByRef ipIndex As Long, ByRef ipLowerIndex As Long, ByRef ipUpperIndex As Long, ByRef ipMessage As String)

    If ipIndex >= ipLowerIndex And ipIndex <= ipUpperIndex Then
        Exit Sub
    End If

    Err.Raise 17 + vbObjectError, _
        ipMessage, _
        Fmt.Text("Expecting a value between {0} and {1}. Got {2}.", ipLowerIndex, ipUpperIndex, ipIndex)
    
 End Sub


Public Sub GuardInvalidRangeObject(ByRef myGroupId As e_Group, ByRef ipItem As Variant, ByRef ipMessage As String)

     Select Case myGroupId
    
        Case e_Group.m_string:      Exit Sub
        Case e_Group.m_array:       Exit Sub
        Case e_Group.m_List:        Exit Sub
        Case e_Group.m_Dictionary:  Exit Sub
        Case Else
            Err.Raise 17 + vbObjectError, _
                ipMessage, _
                Fmt.Text("Expecting string, array, list type or dictionary type.  Got {0}", VBA.TypeName(ipItem))
    End Select
    
End Sub

Sub GuardInvalidIndex(ByRef ipIndex As Long, ByRef ipLastIndex As Long, ByRef ipMessage As String)

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
    
    If ArrayInfo.IsNotArray(ipArray) Then
        Err.Raise 17 + vbObjectError, _
            ipMessage, _
            Fmt.Text("Expecting an array.  Got {0}", VBA.TypeName(ipArray))
    End If
        
End Sub

Public Sub GuardEmptyRangeObject(ByRef ipRange As Variant, ByRef ipMessage As String)

    Dim myLen As Long
    Select Case GroupInfo.Id(ipRange)
        Case e_Group.m_string:                  myLen = VBA.Len(ipRange)
        Case e_Group.m_array:                   myLen = ArrayInfo.Count(ipRange)
        Case e_Group.m_List, m_Dictionary:      myLen = ipRange.Count
    End Select
    
    If myLen < 1 Then
        Exit Sub
    End If
    
    Err.Raise 17 + vbObjectError, _
        ipMessage, _
        "Range object is empty"
        
End Sub
