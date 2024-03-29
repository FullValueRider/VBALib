VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHashC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IGNOREModule
'@folder("Set")
'cHashC Steven Laycock July 2023 Modified for a keyless collection class
' derived from
' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
Option Explicit

Private Const DynTakeOver               As Long = 4
Private Const HMul                      As Long = 3727 'fixed HashSlot-amount and Hash-Multiplikator
Private Const MY_TYPENAME               As String = "cHashC"

Private Type HashTableEntry

    Count                                 As Long
    DataIdxsStat(0 To DynTakeOver - 1)    As Long
    DataIdxsDyn()                         As Long
  
End Type


Private W()                             As Integer
Private sAW                             As SAFEARRAY1D

Private mLastExpectedMaxCount           As Long
Private mLastH                          As Long  ' mEnsureUniqueKeys As Boolean,
Private mCount                          As Long
Private mDTUB                           As Long
Private mHashTableSize                  As Long
'Private mCompareMode                    As VbCompareMethod
Private HashTable()                     As HashTableEntry
Private mItems()                         As Variant
'three vars, to handle the deletes
Private mFIndexes()                     As Long
Private mFCount                         As Long
Private mFUB                            As Long


Private Sub Class_Initialize()
    sAW.cDims = 1
    sAW.cbElements = 2
    sAW.cLocks = 1
    sAW.fFeatures = &H11         'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
    BindArray W, VarPtr(sAW)

    'mCompareMode = vbBinaryCompare
    Reinit 16384                                 'at startup we set it up, to behave ideally for up to 16384 Items

    If LWC(97) = 0 Then InitLWC
    
End Sub


Public Sub Reinit(Optional ByVal ExpectedMaxCount As Long)

    If ExpectedMaxCount <= 0 Then
        ExpectedMaxCount = mLastExpectedMaxCount
    End If
    If ExpectedMaxCount < 128 Then
        ExpectedMaxCount = 128
    End If
    mLastExpectedMaxCount = ExpectedMaxCount

    mHashTableSize = 16
    Do Until mHashTableSize * 2 > ExpectedMaxCount
        mHashTableSize = mHashTableSize * 2
    Loop
    If mHashTableSize > 524288 Then
        mHashTableSize = 524288 ' 2^19
    End If
    ReDim HashTable(0 To mHashTableSize - 1)

    mDTUB = mLastExpectedMaxCount
    ReDim mItems(0 To mDTUB) As Variant
    mCount = 0
    mFCount = 0
    mFUB = 16
    ReDim mFIndexes(0 To mFUB)
End Sub


Public Sub Clear()
    Reinit
End Sub


Public Property Get Count() As Long
    Count = mCount
End Property


'Public Property Get StringCompareMode() As VbCompareMethod
'    StringCompareMode = mCompareMode
'End Property
'
'
'Public Property Let StringCompareMode(ByVal RHS As VbCompareMethod)
'    mCompareMode = RHS
'End Property


Public Function Items() As Variant()

    If mCount = 0 Then
        Items = Array()
        Exit Function
    End If

    Dim i As Long
    Dim j As Long
    Dim V() As Variant: ReDim V(0 To mCount - 1)
    For i = 0 To mCount - 1
        Do While IsEmpty(mItems(j))
            j = j + 1
        Loop
        VariantCopy V(i), ByVal VarPtr(mItems(j)): j = j + 1
    Next

    Items = V
    
End Function

Public Function Exists(ByRef ipItem As Variant) As Boolean
    Exists = FindIndex(ipItem) >= 0
End Function


Public Sub Add(ByRef Key As Variant)
    Dim H As Long
    Dim UB As Long
    Dim i As Long
    If mLastH Then
        H = mLastH: mLastH = 0
    Else
        H = -1
        FindIndex Key, H                 'a Value of -1 for H will skip the Index-Search, returning only H
    End If

    'add the new Pair, prolonging the Keys- and Values-arrays
    If mDTUB < mCount Then
        mDTUB = (mDTUB + 16) * 1.3
        ReDim Preserve mItems(0 To mDTUB)
    End If

    If mFCount > 0 Then
        mFCount = mFCount - 1
        i = mFIndexes(mFCount)
    Else
        i = mCount
    End If
    VariantCopyInd ByVal VarPtr(mItems(i)), ByVal VarPtr(Key)

    'add the new DataIndex to the proper Hash-Buckets
    Select Case HashTable(H).Count
        Case Is < DynTakeOver
            HashTable(H).DataIdxsStat(HashTable(H).Count) = i
        Case DynTakeOver
            ReDim Preserve HashTable(H).DataIdxsDyn(DynTakeOver To DynTakeOver + 3)
            HashTable(H).DataIdxsDyn(DynTakeOver) = i
        Case Else
            UB = UBound(HashTable(H).DataIdxsDyn)
            If UB < HashTable(H).Count Then UB = (UB + 3) * 1.3: ReDim Preserve HashTable(H).DataIdxsDyn(DynTakeOver To UB)
            HashTable(H).DataIdxsDyn(HashTable(H).Count) = i
    End Select
    HashTable(H).Count = HashTable(H).Count + 1

    mCount = mCount + 1
End Sub


Public Sub Remove(ByRef ipItem As Variant)
    If mCount = 0 Then Exit Sub
    Dim H As Long
    Dim Idx As Long
    Dim i As Long
    Dim j As Long
    Idx = FindIndex(ipItem, H)
    If Idx < 0 Then
        Err.Raise 5
    End If

    For i = 0 To HashTable(H).Count - 2
        If i < DynTakeOver Then
            If j = 0 Then If HashTable(H).DataIdxsStat(i) = Idx Then j = i + 1
            If j Then
                If j < DynTakeOver Then
                    HashTable(H).DataIdxsStat(i) = HashTable(H).DataIdxsStat(j): j = j + 1
                Else
                    HashTable(H).DataIdxsStat(i) = HashTable(H).DataIdxsDyn(j): j = j + 1
                End If
            End If
        Else
            If j = 0 Then If HashTable(H).DataIdxsDyn(i) = Idx Then j = i + 1
            If j Then HashTable(H).DataIdxsDyn(i) = HashTable(H).DataIdxsDyn(j): j = j + 1
        End If
    Next
    HashTable(H).Count = i

    If mFUB < mFCount Then
        mFUB = mFUB + mFUB
        ReDim Preserve mFIndexes(0 To mFUB)
    End If
    
    mFIndexes(mFCount) = Idx
    mFCount = mFCount + 1

    mItems(Idx) = Empty
    mCount = mCount - 1
End Sub


Public Sub RemoveByIndex(ByVal IndexZeroBased As Long)
    Remove ItemByIndex(IndexZeroBased)
End Sub


Public Function IndexByItem(ByRef Key As Variant) As Long
    Dim Idx As Long
    Idx = FindIndex(Key)
    If Idx >= 0 And mFCount > 0 Then AdjustIndex Idx, True
    IndexByItem = Idx
End Function


Public Property Get ItemByIndex(ByVal IndexZeroBased As Long) As Variant
    If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
    If mFCount Then AdjustIndex IndexZeroBased
    VariantCopy ItemByIndex, ByVal VarPtr(mItems(IndexZeroBased))
End Property


'Public Property Get ItemByIndex(ByVal IndexZeroBased As Long) As Variant
'  If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
'  If mFCount Then AdjustIndex IndexZeroBased
'  VariantCopy ItemByIndex, ByVal VarPtr(mValues(IndexZeroBased))
'End Property
'
'Public Property Let ItemByIndex(ByVal IndexZeroBased As Long, ByVal RHS As Variant)
'  If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
'  If mFCount Then AdjustIndex IndexZeroBased
'  VariantCopyInd ByVal VarPtr(mValues(IndexZeroBased)), ByVal VarPtr(RHS)
'End Property
'
'Public Property Set ItemByIndex(ByVal IndexZeroBased As Long, ByVal RHS As Variant)
'  If IndexZeroBased < 0 Or IndexZeroBased >= mCount Then Err.Raise 9
'  If mFCount Then AdjustIndex IndexZeroBased
'  VariantCopyInd ByVal VarPtr(mValues(IndexZeroBased)), ByVal VarPtr(RHS)
'End Property

'Public Property Get Item(ByRef Key As Variant) As Variant
'Dim Index As Long:  Index = FindIndex(Key)
' If Index >= 0 Then VariantCopy Item, ByVal VarPtr(mValues(Index))
'End Property
'Public Property Let Item(ByVal Key As Variant, ByVal RHS As Variant)
'Dim Index As Long:  Index = FindIndex(Key, mLastH)
' If Index = -1 Then Add Key, RHS Else VariantCopyInd ByVal VarPtr(mValues(Index)), ByVal VarPtr(RHS)
' mLastH = 0
'End Property
'Public Property Set Item(ByVal Key As Variant, ByVal RHS As Variant)
'Dim Index As Long:  Index = FindIndex(Key, mLastH)
' If Index = -1 Then Add Key, RHS Else VariantCopyInd ByVal VarPtr(mValues(Index)), ByVal VarPtr(RHS)
' mLastH = 0
'End Property

Private Function FindIndex(ByRef Key As Variant, Optional ByRef H As Long) As Long 'return -1, when no Key can be found
    Dim i As Long
    Dim C As Variant
    Dim D As Double
    Dim L As LongPtr
    Dim F As Single
    Dim VT As VbVarType
    Dim HTUB As Long
    ' variables can be declared as decimal using the @ type character but cannot using 'as Decimal'
    ' hence we declare c as variant then converyt to decimalusing vba.cdec
    C = VBA.CDec(C)
    HTUB = mHashTableSize - 1
    FindIndex = H
    H = HTUB                                     'init the HashValue (all bits to 1)

    VT = VarType(Key)

    'Dim mySaw As SAFEARRAY1D


    Select Case VT
        Case vbString
            sAW.cElements1D = VBA.Len(Key)
            sAW.pvData = StrPtr(Key)
            'If mCompareMode = 0 Then
                For i = 0 To sAW.cElements1D - 1
                    H = (H + W(i)) * HMul And HTUB
                Next
                If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

                For i = 0 To HashTable(H).Count - 1
                    If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
                    If VarType(mItems(FindIndex)) = VT Then If Key = mItems(FindIndex) Then Exit Function
                Next
'            Else
'                For i = 0 To sAW.cElements1D - 1: H = (H + LWC(W(i))) * HMul And HTUB: Next
'                If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation
'
'                For i = 0 To HashTable(H).Count - 1
'                    If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
'                    If VarType(mItems(FindIndex)) = VT Then If StrComp(Key, mItems(FindIndex), mCompareMode) = 0 Then Exit Function
'                Next
'            End If

        Case vbObject
            L = ObjPtr(Key): sAW.cElements1D = 2: sAW.pvData = VarPtr(L)
            H = (H + W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + W(1)) * HMul And HTUB
            If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

            For i = 0 To HashTable(H).Count - 1
                If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
                If VarType(mItems(FindIndex)) = VT Then If Key Is mItems(FindIndex) Then Exit Function
            Next

        Case vbCurrency
            C = Key: sAW.cElements1D = 4: sAW.pvData = VarPtr(C)
            H = (H + W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + W(1)) * HMul And HTUB
            H = (H + W(2)) * HMul And HTUB
            H = (H + W(3)) * HMul And HTUB
            If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

            For i = 0 To HashTable(H).Count - 1
                If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
                If VarType(mItems(FindIndex)) = VT Then If C = mItems(FindIndex) Then Exit Function
            Next

        Case vbLong, vbInteger, vbByte
            L = Key: sAW.cElements1D = 2: sAW.pvData = VarPtr(L)
            H = (H + W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + W(1)) * HMul And HTUB
            If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

            For i = 0 To HashTable(H).Count - 1
                If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
                Select Case VarType(mItems(FindIndex))
                    Case vbLong, vbInteger, vbByte: If L = mItems(FindIndex) Then Exit Function
                End Select
            Next

        Case vbDouble
            D = Key: sAW.cElements1D = 4: sAW.pvData = VarPtr(D)
            H = (H + W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + W(1)) * HMul And HTUB
            H = (H + W(2)) * HMul And HTUB
            H = (H + W(3)) * HMul And HTUB
            If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

            For i = 0 To HashTable(H).Count - 1
                If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
                If VarType(mItems(FindIndex)) = VT Then If D = mItems(FindIndex) Then Exit Function
            Next

        Case vbDate
            D = Key: sAW.cElements1D = 4: sAW.pvData = VarPtr(D)
            H = (H + W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + W(1)) * HMul And HTUB
            H = (H + W(2)) * HMul And HTUB
            H = (H + W(3)) * HMul And HTUB
            If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

            For i = 0 To HashTable(H).Count - 1
                If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
                If VarType(mItems(FindIndex)) = VT Then If D = mItems(FindIndex) Then Exit Function
            Next

        Case vbSingle
            F = Key: sAW.cElements1D = 2: sAW.pvData = VarPtr(F)
            H = (H + W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + W(1)) * HMul And HTUB
            If FindIndex = -1 Then Exit Function 'it's a "Hash-Only" Calculation

            For i = 0 To HashTable(H).Count - 1
                If i < DynTakeOver Then FindIndex = HashTable(H).DataIdxsStat(i) Else FindIndex = HashTable(H).DataIdxsDyn(i)
                If VarType(mItems(FindIndex)) = VT Then If F = mItems(FindIndex) Then Exit Function
            Next
    End Select

    FindIndex = -1
End Function


'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub AdjustIndex(ByRef Idx As Long, Optional ByVal Inverse As Boolean)
    Dim i As Long
    Dim Incr As Long
    Incr = IIf(Inverse, -1, 1)
    'If Inverse Then Incr = -1 Else Incr = 1
    For i = 0 To mFCount - 1
        If mFIndexes(i) <= Idx Then Idx = Idx + Incr
    Next
End Sub


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property


'Friend Sub CheckHashDistribution()
'Dim i As Long, Count As Long, cc As Long, Min As Long, Max As Long
'  Min = &H7FFFFFFF
'  For i = 0 To UBound(HashTable)
'    Count = HashTable(i).Count
'    If Count Then
'      If Min > Count Then Min = Count
'      If Max < Count Then Max = Count
'      cc = cc + 1
'    End If
'  Next
'  Debug.Print "Distribution of"; mCount; "entries over a HashTable with"; UBound(HashTable) + 1; "slots:"
'  Debug.Print "Used-HashSlots:"; cc
'  Debug.Print "Min-Entries:"; Min
'  Debug.Print "Max-Entries:"; Max
'End Sub
