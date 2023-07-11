Attribute VB_Name = "Sorters"
Option Explicit

Public Sub ShakerSortArray(ByRef iopArray As Variant)
    ' from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2912324#post2912324
    ' The shaker sort is used because it is the highest rated sort which is stable and inplace and which does not use recursion
    If Not VBA.IsArray(iopArray) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.Sorters.ShakerSortArray", _
        Fmt.Text("Expecting array. Got {0}.", VBA.TypeName(iopArray))
        
    End If
    
    If ArrayInfo.LacksItems(iopArray) Then
        Exit Sub
    End If
    
    If ArrayInfo.IsNotArray(iopArray, e_ArrayType.m_ListArray) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Sorters.ShakerSortArray", _
            Fmt.Text("Expecting array with one dimensions. Got {0} dimensions", ArrayInfo.Ranks(iopArray))
    End If
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = LBound(iopArray)
    iMax = UBound(iopArray)
    i = (iMax - iMin) \ 2 + iMin
    Do While i > iMin
        j = i
        Do While j > iMin
            For k = iMin To i - j
                If iopArray(k) > iopArray(k + j) Then
                    If VBA.IsObject(iopArray(k)) Then
                        Set varSwap = iopArray(k)
                    Else
                        varSwap = iopArray(k)
                    End If
                    If VBA.IsObject(iopArray(k + j)) Then
                        Set iopArray(k) = iopArray(k + j)
                    Else
                        iopArray(k) = iopArray(k + j)
                    End If
                    If VBA.IsObject(varSwap) Then
                        Set iopArray(k + j) = varSwap
                    Else
                        iopArray(k + j) = varSwap
                    End If
                End If
            Next
            j = j \ 2
        Loop
        i = i \ 2
    Loop
    iMax = iMax - 1
    Do
        blnSwapped = False
        For i = iMin To iMax
            If iopArray(i) > iopArray(i + 1) Then
                If VBA.IsObject(iopArray(i)) Then
                    Set varSwap = iopArray(i)
                Else
                    varSwap = iopArray(i)
                End If
                If VBA.IsObject(iopArray(i + 1)) Then
                    Set iopArray(i) = iopArray(i + 1)
                Else
                    iopArray(i) = iopArray(i + 1)
                End If
                If VBA.IsObject(varSwap) Then
                    Set iopArray(i + 1) = varSwap
                Else
                    iopArray(i + 1) = varSwap
                End If
                blnSwapped = True
            End If
        Next i
        If blnSwapped Then
            blnSwapped = False
            iMax = iMax - 1
            For i = iMax To iMin Step -1
                If iopArray(i) > iopArray(i + 1) Then
                    If VBA.IsObject(iopArray(i)) Then
                    Set varSwap = iopArray(i)
                Else
                    varSwap = iopArray(i)
                End If
                If VBA.IsObject(iopArray(i + 1)) Then
                    Set iopArray(i) = iopArray(i + 1)
                Else
                    iopArray(i) = iopArray(i + 1)
                End If
                If VBA.IsObject(varSwap) Then
                    Set iopArray(i + 1) = varSwap
                Else
                    iopArray(i + 1) = varSwap
                End If
                    blnSwapped = True
                End If
            Next i
            iMin = iMin + 1
        End If
    Loop Until Not blnSwapped
End Sub

Public Sub ShakerSortSeq(ByRef iopS As Object)
    ' from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2912324#post2912324
    ' The shaker sort is used because it is the highest rated sort which is stable and inplace and which does not use recursion
    
    If VBA.Left$(VBA.TypeName(iopS), 3) <> "Seq" Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Sorters.ShakerSortSeq", _
            Fmt.Text("Expecting a Seq.  Got {0}", VBA.TypeName(iopS))
    End If
    
    If iopS.Count < 1 Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = iopS.FirstIndex
    iMax = iopS.LastIndex
    i = (iMax - iMin) \ 2 + iMin
    Do While i > iMin
        j = i
        Do While j > iMin
            For k = iMin To i - j
                If iopS.Item(k) > iopS.Item(k + j) Then
                    If VBA.IsObject(iopS.Item(k)) Then
                        Set varSwap = iopS.Item(k)
                    Else
                        varSwap = iopS.Item(k)
                    End If
                    If VBA.IsObject(iopS.Item(k + j)) Then
                        Set iopS.Item(k) = iopS.Item(k + j)
                    Else
                        iopS.Item(k) = iopS.Item(k + j)
                    End If
                    If VBA.IsObject(varSwap) Then
                        Set iopS.Item(k + j) = varSwap
                    Else
                        iopS.Item(k + j) = varSwap
                    End If
                End If
            Next
            j = j \ 2
        Loop
        i = i \ 2
    Loop
    iMax = iMax - 1
    Do
        blnSwapped = False
        For i = iMin To iMax
            If iopS.Item(i) > iopS.Item(i + 1) Then
                If VBA.IsObject(iopS.Item(i)) Then
                    Set varSwap = iopS.Item(i)
                Else
                    varSwap = iopS.Item(i)
                End If
                If VBA.IsObject(iopS.Item(i + 1)) Then
                    Set iopS.Item(i) = iopS.Item(i + 1)
                Else
                    iopS.Item(i) = iopS.Item(i + 1)
                End If
                If VBA.IsObject(varSwap) Then
                    Set iopS.Item(i + 1) = varSwap
                Else
                    iopS.Item(i + 1) = varSwap
                End If
                blnSwapped = True
            End If
        Next i
        If blnSwapped Then
            blnSwapped = False
            iMax = iMax - 1
            For i = iMax To iMin Step -1
                If iopS.Item(i) > iopS.Item(i + 1) Then
                    If VBA.IsObject(iopS.Item(i)) Then
                    Set varSwap = iopS.Item(i)
                Else
                    varSwap = iopS.Item(i)
                End If
                If VBA.IsObject(iopS.Item(i + 1)) Then
                    Set iopS.Item(i) = iopS.Item(i + 1)
                Else
                    iopS.Item(i) = iopS.Item(i + 1)
                End If
                If VBA.IsObject(varSwap) Then
                    Set iopS.Item(i + 1) = varSwap
                Else
                    iopS.Item(i + 1) = varSwap
                End If
                    blnSwapped = True
                End If
            Next i
            iMin = iMin + 1
        End If
    Loop Until Not blnSwapped
End Sub
