Attribute VB_Name = "Sorters"
'@Folder("Helpers")
Option Explicit


Public Sub ShakerSortArrayByIndex(ByRef iopArray As Variant)
    ' from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2912324#post2912324
    ' The shaker sort is used because it is the highest rated sort which is stable and inplace and which does not use recursion
    If Not VBA.IsArray(iopArray) Then
        Err.Raise 17 + vbObjectError, _
        "Sorters.ShakerSortArray", _
        Fmt.Text("Expecting array. Got {0}.", VBA.TypeName(iopArray))
        
    End If
    
    If ArrayOp.LacksItems(iopArray) Then
        Exit Sub
    End If
    
    If ArrayOp.IsNotArray(iopArray, e_ArrayType.m_ListArray) Then
        Err.Raise 17 + vbObjectError, _
        "Sorters.ShakerSortArray", _
        Fmt.Text("Expecting array with one dimensions. Got {0} dimensions", ArrayOp.Ranks(iopArray))
    End If
    
    Dim i As Long
    Dim j As Long
    Dim K As Long
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
            For K = iMin To i - j
                If Comparers.MT(iopArray(K), iopArray(K + j)) Then
                    If VBA.IsObject(iopArray(K)) Then
                        Set varSwap = iopArray(K)
                    Else
                        varSwap = iopArray(K)
                    End If
                    If VBA.IsObject(iopArray(K + j)) Then
                        Set iopArray(K) = iopArray(K + j)
                    Else
                        iopArray(K) = iopArray(K + j)
                    End If
                    If VBA.IsObject(varSwap) Then
                        Set iopArray(K + j) = varSwap
                    Else
                        iopArray(K + j) = varSwap
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
            If Comparers.MT(iopArray(i), iopArray(i + 1)) Then
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
                If Comparers.MT(iopArray(i), iopArray(i + 1)) Then
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

Public Sub ShakerSortArrayByItemOfIndex(ByRef iopArray As Variant)
    ' from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2912324#post2912324
    ' The shaker sort is used because it is the highest rated sort which is stable and inplace and which does not use recursion
    If Not VBA.IsArray(iopArray) Then
        Err.Raise 17 + vbObjectError, _
        "Sorters.ShakerSortArray", _
        Fmt.Text("Expecting array. Got {0}.", VBA.TypeName(iopArray))
        
    End If
    
    If ArrayOp.LacksItems(iopArray) Then
        Exit Sub
    End If
    
    If ArrayOp.IsNotArray(iopArray, e_ArrayType.m_ListArray) Then
        Err.Raise 17 + vbObjectError, _
        "Sorters.ShakerSortArray", _
        Fmt.Text("Expecting array with one dimensions. Got {0} dimensions", ArrayOp.Ranks(iopArray))
    End If
    
    Dim i As Long
    Dim j As Long
    Dim K As Long
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
            For K = iMin To i - j
                If Comparers.MT(iopArray(K).Item, iopArray(K + j).Item) Then
                    
                    Set varSwap = iopArray(K)
                    Set iopArray(K) = iopArray(K + j)
                    Set iopArray(K + j) = varSwap
                   
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
            If Comparers.MT(iopArray(i).Item, iopArray(i + 1).Item) Then
                
                Set varSwap = iopArray(i)
                Set iopArray(i) = iopArray(i + 1)
                Set iopArray(i + 1) = varSwap
                blnSwapped = True
                
            End If
        Next i
        If blnSwapped Then
            blnSwapped = False
            iMax = iMax - 1
            For i = iMax To iMin Step -1
                If Comparers.MT(iopArray(i).Item, iopArray(i + 1).Item) Then
                    
                    Set varSwap = iopArray(i)
                    Set iopArray(i) = iopArray(i + 1)
                    Set iopArray(i + 1) = varSwap
                    blnSwapped = True
                    
                End If
            Next i
            iMin = iMin + 1
        End If
    Loop Until Not blnSwapped
End Sub


Public Sub ShakerSortByItem(ByVal iopS As Object)
    ' from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2912324#post2912324
    ' The shaker sort is used because it is the highest rated sort which is stable and inplace and which does not use recursion
    
    If VBA.Left$(VBA.TypeName(iopS), 3) <> "Seq" Then
        Err.Raise 17 + vbObjectError, _
        "Sorters.ShakerSortByItem", _
        Fmt.Text("Expecting a Seq.  Got {0}", VBA.TypeName(iopS))
    End If
    
    If iopS.Count < 1 Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = iopS.FirstIndex
    iMax = iopS.Lastindex
    i = (iMax - iMin) \ 2 + iMin
    Do While i > iMin
        j = i
        Do While j > iMin
            For K = iMin To i - j
                If Comparers.MT(iopS.Item(K), iopS.Item(K + j)) Then
                    If VBA.IsObject(iopS.Item(K)) Then
                        Set varSwap = iopS.Item(K)
                    Else
                        varSwap = iopS.Item(K)
                    End If
                    If VBA.IsObject(iopS.Item(K + j)) Then
                        Set iopS.Item(K) = iopS.Item(K + j)
                    Else
                        iopS.Item(K) = iopS.Item(K + j)
                    End If
                    If VBA.IsObject(varSwap) Then
                        Set iopS.Item(K + j) = varSwap
                    Else
                        iopS.Item(K + j) = varSwap
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
            If Comparers.MT(iopS.Item(i), iopS.Item(i + 1)) Then
                'Swap iopS.Item(i), iopS.Item(i + 1)
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
                If Comparers.MT(iopS.Item(i), iopS.Item(i + 1)) Then
                    'Swap iopS.Item(i), iopS.Item(i + 1)
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



Public Sub ShakerSortByItemOfItem(ByVal iopS As Object)
    ' from https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2912324#post2912324
    ' The shaker sort is used because it is the highest rated sort which is stable and inplace and which does not use recursion
    
    If VBA.Left$(VBA.TypeName(iopS), 3) <> "Seq" Then
        Err.Raise 17 + vbObjectError, _
        "Sorters.ShakerSortByItem", _
        Fmt.Text("Expecting a Seq.  Got {0}", VBA.TypeName(iopS))
    End If
    
    If iopS.Count < 1 Then
        Exit Sub
    End If
    
    Dim i As Long
    Dim j As Long
    Dim K As Long
    Dim iMin As Long
    Dim iMax As Long
    Dim varSwap As Variant
    Dim blnSwapped As Boolean
    
    iMin = iopS.FirstIndex
    iMax = iopS.Lastindex
    i = (iMax - iMin) \ 2 + iMin
    Do While i > iMin
        j = i
        Do While j > iMin
            For K = iMin To i - j
                If Comparers.MT(iopS.Item(K).Item, iopS.Item(K + j).Item) Then
                    
                    Set varSwap = iopS.Item(K)
                    Set iopS.Item(K) = iopS.Item(K + j)
                    Set iopS.Item(K + j) = varSwap
                    iopS.Item(K + j) = varSwap
                    
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
            If Comparers.MT(iopS.Item(i).Item, iopS.Item(i + 1).Item) Then
                
                 Set varSwap = iopS.Item(i)
                 Set iopS.Item(i) = iopS.Item(i + 1)
                 Set iopS.Item(i + 1) = varSwap
                blnSwapped = True
                
            End If
        Next i
        If blnSwapped Then
            blnSwapped = False
            iMax = iMax - 1
            For i = iMax To iMin Step -1
                If Comparers.EQ(iopS.Item(i).Item, iopS.Item(i + 1).Item) Then
                    
                    Set varSwap = iopS.Item(i)
                    Set iopS.Item(i) = iopS.Item(i + 1)
                    Set iopS.Item(i + 1) = varSwap
                    blnSwapped = True
                    
                End If
            Next i
            iMin = iMin + 1
        End If
    Loop Until Not blnSwapped
End Sub




