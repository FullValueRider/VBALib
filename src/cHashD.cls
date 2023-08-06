VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHashD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========

' Based on cHashD from https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&s=f049d882b4fc4d27f4bc2f9bc383a53e
' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
' Reformatted and Rubberducked by Steven Laycock July 2023
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========

Option Explicit

Private Const NoEntry                       As Currency = -98765432123456@
'fixed HashSlot-amount and Hash-Multiplikator
Private Const DynTakeOver                   As Long = 4
Private Const HMul                          As Long = 3727

Private Type HashTableEntry
    Count                                     As Long
    DataIdxsStat(0 To DynTakeOver - 1)        As Long
    DataIdxsDyn()                             As Long
End Type

Private Type State
     W()                                    As Integer
     saW                                    As SAFEARRAY1D
     
     LastExpectedMaxCount                   As Long
     EnsureUniqueKeys                       As Boolean
     LastH                                  As Long
     Count                                  As Long
     DTUB                                   As Long
     HashTableSize                          As Long
     CompareMode                            As VbCompareMethod
     HashTable()                            As HashTableEntry
    'both pairing-arrays are of type Variant
     Keys()                                 As Variant
     Values()                               As Variant
    'three vars, to handle the deletes
     FIndexes()                             As Long
     FCount                                 As Long
     FUB                                    As Long
     
End Type

Private s                                   As State


Private Sub Class_Initialize()

    s.saW.cDims = 1
    s.saW.cbElements = 2
    s.saW.cLocks = 1
    s.saW.fFeatures = &H11                         'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
    BindArray s.W, VarPtr(s.saW)
    
    s.CompareMode = vbBinaryCompare
    ReInit 16384                                 'at startup we set it up, to behave ideally for up to 16384 Items
    
    If LWC(97) = 0 Then
        InitLWC
    End If
    
End Sub


Public Sub ReInit(Optional ByVal ExpectedMaxCount As Long, Optional ByVal EnsureUniqueKeys As Boolean)

    If ExpectedMaxCount <= 0 Then
        ExpectedMaxCount = s.LastExpectedMaxCount
    End If
    
    If ExpectedMaxCount < 128 Then
        ExpectedMaxCount = 128
    End If
    
    s.LastExpectedMaxCount = ExpectedMaxCount
    s.EnsureUniqueKeys = EnsureUniqueKeys
    
    s.HashTableSize = 16
    Do Until s.HashTableSize * 2 > ExpectedMaxCount
        s.HashTableSize = s.HashTableSize * 2
    Loop
    
    If s.HashTableSize > 524288 Then
        s.HashTableSize = 524288
    End If
    
    ReDim s.HashTable(0 To s.HashTableSize - 1)
    
    s.DTUB = s.LastExpectedMaxCount
    ReDim s.Keys(0 To s.DTUB)
    ReDim s.Values(0 To s.DTUB)
    s.Count = 0
    s.FCount = 0
    s.FUB = 16
    ReDim s.FIndexes(0 To s.FUB)
    
End Sub


Public Sub Clear()
    ReInit
End Sub


Public Property Get Count() As Long
    Count = s.Count
End Property


Public Property Get HasUniqueKeys() As Boolean
    HasUniqueKeys = s.EnsureUniqueKeys
End Property


Public Property Get StringCompareMode() As VbCompareMethod
    StringCompareMode = s.CompareMode
End Property


Public Property Let StringCompareMode(ByVal RHS As VbCompareMethod)
    s.CompareMode = RHS
End Property


Public Function Keys() As Variant()

    If s.Count = 0 Then
        Keys = Array()
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Long
    Dim v() As Variant
    ReDim v(0 To s.Count - 1)
    
    For i = 0 To s.Count - 1
    
        Do While s.Values(j) = NoEntry
            j = j + 1
        Loop
        
        VariantCopy v(i), ByVal VarPtr(s.Keys(j))
        j = j + 1
        
    Next
    
    Keys = v
    
End Function


Public Function Items() As Variant()

    If s.Count = 0 Then
        Items = Array()
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Long
    
    Dim v() As Variant
    ReDim v(0 To s.Count - 1)
    
    For i = 0 To s.Count - 1
        Do While s.Values(j) = NoEntry
            j = j + 1
        Loop
    
        VariantCopy v(i), ByVal VarPtr(s.Values(j))
        j = j + 1
        
    Next
    
    Items = v
    
End Function


Public Function Pairs() As Variant()             'hand-out the Key-Values in a Pairs-Array

    If s.Count = 0 Then
        Pairs = Array()
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Long
    
    Dim v() As Variant
    ReDim v(0 To s.Count - 1)
    
    For i = 0 To s.Count - 1
    
        Do While s.Values(j) = NoEntry
            j = j + 1:
        Loop
        
        v(i) = Array(s.Keys(j), s.Values(j))
        j = j + 1
        
    Next
    
    Pairs = v
    
End Function


Public Function Exists(ByRef Key As Variant) As Boolean
    Exists = FindIndex(Key) >= 0
End Function


Public Function HoldsKey(ByRef ipKey As Variant) As Boolean
    HoldsKey = FindIndex(ipKey) >= 0
End Function


Public Function LacksKey(ByRef ipKey As Variant) As Boolean
    LacksKey = FindIndex(ipKey) < 0
End Function


Public Sub Add(ByRef Key As Variant, ByRef Item As Variant)

    Dim H As Long
    Dim UB As Long
    Dim i As Long
    
    If s.LastH Then
        H = s.LastH
        s.LastH = 0
    ElseIf s.EnsureUniqueKeys Then
        If FindIndex(Key, H) >= 0 Then
            Err.Raise 457
        End If
    Else
        H = -1
        FindIndex Key, H                         'a Value of -1 for H will skip the Index-Search, returning only H
    End If
    
    'add the new Pair, prolonging the Keys- and Values-arrays
    If s.DTUB < s.Count Then
        s.DTUB = (s.DTUB + 16) * 1.3
        ReDim Preserve s.Keys(0 To s.DTUB)
        ReDim Preserve s.Values(0 To s.DTUB)
    End If
    
    If s.FCount > 0 Then
        s.FCount = s.FCount - 1
        i = s.FIndexes(s.FCount)
    Else
        i = s.Count
    End If
    
    VariantCopyInd ByVal VarPtr(s.Values(i)), ByVal VarPtr(Item)
    VariantCopyInd ByVal VarPtr(s.Keys(i)), ByVal VarPtr(Key)
    
    'add the new DataIndex to the proper Hash-Buckets
    Select Case s.HashTable(H).Count
    
        Case Is < DynTakeOver
        
            s.HashTable(H).DataIdxsStat(s.HashTable(H).Count) = i
            
        Case DynTakeOver
        
            ReDim Preserve s.HashTable(H).DataIdxsDyn(DynTakeOver To DynTakeOver + 3)
            s.HashTable(H).DataIdxsDyn(DynTakeOver) = i
            
        Case Else
        
            UB = UBound(s.HashTable(H).DataIdxsDyn)
            
            If UB < s.HashTable(H).Count Then
                UB = (UB + 3) * 1.3
                ReDim Preserve s.HashTable(H).DataIdxsDyn(DynTakeOver To UB)
            End If
            
            s.HashTable(H).DataIdxsDyn(s.HashTable(H).Count) = i
            
    End Select
    
    s.HashTable(H).Count = s.HashTable(H).Count + 1
    
    s.Count = s.Count + 1
    
End Sub


Public Sub Remove(ByRef Key As Variant)

    If s.Count = 0 Then
        Exit Sub
    End If
    
    Dim H As Long
    Dim Idx As Long
    Dim i As Long
    Dim j As Long
    
    Idx = FindIndex(Key, H)
    If Idx < 0 Then
        Err.Raise 5
    End If
    
    For i = 0 To s.HashTable(H).Count - 2
        If i < DynTakeOver Then
        
            If j = 0 Then
                If s.HashTable(H).DataIdxsStat(i) = Idx Then
                    j = i + 1
                End If
            End If
            
            If j Then
                If j < DynTakeOver Then
                    s.HashTable(H).DataIdxsStat(i) = s.HashTable(H).DataIdxsStat(j)
                    j = j + 1
                Else
                    s.HashTable(H).DataIdxsStat(i) = s.HashTable(H).DataIdxsDyn(j)
                    j = j + 1
                End If
            End If
        Else
        
            If j = 0 Then
                If s.HashTable(H).DataIdxsDyn(i) = Idx Then
                    j = i + 1
                End If
            End If
            
            If j Then
                s.HashTable(H).DataIdxsDyn(i) = s.HashTable(H).DataIdxsDyn(j)
                j = j + 1
            End If
            
        End If
    Next
    
    s.HashTable(H).Count = i
    
    If s.FUB < s.FCount Then
        s.FUB = s.FUB + s.FUB
        ReDim Preserve s.FIndexes(0 To s.FUB)
    End If
    
    s.FIndexes(s.FCount) = Idx
    s.FCount = s.FCount + 1
    
    s.Keys(Idx) = NoEntry
    s.Values(Idx) = NoEntry
    s.Count = s.Count - 1
    
End Sub


Public Sub RemoveByIndex(ByVal IndexZeroBased As Long)
    Remove KeyByIndex(IndexZeroBased)
End Sub


Public Function IndexByKey(ByRef Key As Variant) As Long

    Dim Idx As Long
    Idx = FindIndex(Key)
    
    If Idx >= 0 And s.FCount > 0 Then
        AdjustIndex Idx, True
    End If
    
    IndexByKey = Idx
    
End Function


Public Property Get KeyByIndex(ByVal IndexZeroBased As Long) As Variant

    If IndexZeroBased < 0 Or IndexZeroBased >= s.Count Then
        Err.Raise 9
    End If
    
    If s.FCount Then
        AdjustIndex IndexZeroBased
    End If
    
    VariantCopy KeyByIndex, ByVal VarPtr(s.Keys(IndexZeroBased))
    
End Property


Public Property Get ItemByIndex(ByVal IndexZeroBased As Long) As Variant

    If IndexZeroBased < 0 Or IndexZeroBased >= s.Count Then
        Err.Raise 9
    End If
    
    If s.FCount Then
        AdjustIndex IndexZeroBased
    End If
    
    VariantCopy ItemByIndex, ByVal VarPtr(s.Values(IndexZeroBased))
    
End Property


Public Property Let ItemByIndex(ByVal IndexZeroBased As Long, ByVal RHS As Variant)

    If IndexZeroBased < 0 Or IndexZeroBased >= s.Count Then
        Err.Raise 9
    End If
    
    If s.FCount Then
        AdjustIndex IndexZeroBased
    End If
    
    VariantCopyInd ByVal VarPtr(s.Values(IndexZeroBased)), ByVal VarPtr(RHS)
    
End Property


Public Property Set ItemByIndex(ByVal IndexZeroBased As Long, ByVal RHS As Variant)

    If IndexZeroBased < 0 Or IndexZeroBased >= s.Count Then
        Err.Raise 9
    End If
    
    If s.FCount Then
        AdjustIndex IndexZeroBased
    End If
    
    VariantCopyInd ByVal VarPtr(s.Values(IndexZeroBased)), ByVal VarPtr(RHS)
    
End Property


Public Property Get Item(ByRef Key As Variant) As Variant

    Dim Index As Long
    Index = FindIndex(Key)
    
    If Index >= 0 Then
        VariantCopy Item, ByVal VarPtr(s.Values(Index))
    End If
    
End Property


Public Property Let Item(ByRef Key As Variant, ByVal RHS As Variant)

    Dim Index As Long
    Index = FindIndex(Key, s.LastH)
    
    If Index = -1 Then
        Add Key, RHS
    Else
        VariantCopyInd ByVal VarPtr(s.Values(Index)), ByVal VarPtr(RHS)
    End If
    
    s.LastH = 0
    
End Property


Public Property Set Item(ByRef Key As Variant, ByVal RHS As Variant)

    Dim Index As Long
    Index = FindIndex(Key, s.LastH)
    
    If Index = -1 Then
        Add Key, RHS
    Else
        VariantCopyInd ByVal VarPtr(s.Values(Index)), ByVal VarPtr(RHS)
    End If
    
    s.LastH = 0
    
End Property


Private Function FindIndex(ByRef Key As Variant, Optional ByRef H As Long) As Long 'return -1, when no Key can be found

    Dim i As Long
    Dim D As Double
    Dim L As LongPtr
    Dim F As Single
    Dim VT As VbVarType
    Dim HTUB As Long
     
    Dim C As Variant
    C = VBA.CDec(C)
     
    HTUB = s.HashTableSize - 1
    FindIndex = H
    H = HTUB                                     'init the HashValue (all bits to 1)
     
    VT = VarType(Key)
    Select Case VT
        Case vbString
        
            s.saW.cElements1D = Len(Key)
            s.saW.pvData = StrPtr(Key)
            
            If s.CompareMode = 0 Then
            
                For i = 0 To s.saW.cElements1D - 1
                    H = (H + s.W(i)) * HMul And HTUB
                Next
                
                If FindIndex = -1 Then
                    Exit Function 'it's a "Hash-Only" Calculation
                End If
            
                For i = 0 To s.HashTable(H).Count - 1
                
                    If i < DynTakeOver Then
                        FindIndex = s.HashTable(H).DataIdxsStat(i)
                    Else
                        FindIndex = s.HashTable(H).DataIdxsDyn(i)
                    End If
                    
                    If VarType(s.Keys(FindIndex)) = VT Then
                        If Key = s.Keys(FindIndex) Then
                            Exit Function
                        End If
                    End If
                    
                Next
                
            Else
            
                For i = 0 To s.saW.cElements1D - 1
                    H = (H + LWC(s.W(i))) * HMul And HTUB
                Next
                
                If FindIndex = -1 Then
                    Exit Function 'it's a "Hash-Only" Calculation
                End If
            
                For i = 0 To s.HashTable(H).Count - 1
                
                    If i < DynTakeOver Then
                        FindIndex = s.HashTable(H).DataIdxsStat(i)
                    Else
                        FindIndex = s.HashTable(H).DataIdxsDyn(i)
                    End If
                    
                    If VarType(s.Keys(FindIndex)) = VT Then
                        If StrComp(Key, s.Keys(FindIndex), s.CompareMode) = 0 Then
                            Exit Function
                        End If
                    End If
                    
                Next
                
            End If
         
        Case vbObject
        
            L = ObjPtr(Key)
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(L)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If Key Is s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
         
        Case vbCurrency
        
            C = Key
            s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(C)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            H = (H + s.W(2)) * HMul And HTUB
            H = (H + s.W(3)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If C = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
         
        Case vbLong, vbInteger, vbByte
        
            L = Key
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(L)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
         
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                Select Case VarType(s.Keys(FindIndex))
                    Case vbLong, vbInteger, vbByte
                        If L = s.Keys(FindIndex) Then
                            Exit Function
                        End If
                End Select
                
            Next
    
        Case vbDouble
        
            D = Key: s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(D)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            H = (H + s.W(2)) * HMul And HTUB
            H = (H + s.W(3)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If D = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
       
        Case vbDate
        
            D = Key
            s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(D)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            H = (H + s.W(2)) * HMul And HTUB
            H = (H + s.W(3)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If D = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
         
        Case vbSingle
            F = Key
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(F)
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If F = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
    End Select
     
    FindIndex = -1
    
End Function


'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub AdjustIndex(ByRef Idx As Long, Optional ByVal Inverse As Boolean)

    Dim i As Long
    Dim Incr As Long
    
    'If Inverse Then Incr = -1 Else Incr = 1
    Incr = IIf(Inverse, -1, 1)
    
    For i = 0 To s.FCount - 1
        If s.FIndexes(i) <= Idx Then
            Idx = Idx + Incr
        End If
    Next
End Sub


Friend Sub CheckHashDistribution()
    Dim i As Long
    Dim Count As Long
    Dim cc As Long
    Dim Min As Long
    Dim Max As Long
    Min = &H7FFFFFFF
    For i = 0 To UBound(s.HashTable)
        Count = s.HashTable(i).Count
        If Count Then
            If Min > Count Then Min = Count
            If Max < Count Then Max = Count
            cc = cc + 1
        End If
    Next
    Debug.Print "Distribution of"; s.Count; "entries over a HashTable with"; UBound(s.HashTable) + 1; "slots:"
    Debug.Print "Used-HashSlots:"; cc
    Debug.Print "Min-Entries:"; Min
    Debug.Print "Max-Entries:"; Max
End Sub

