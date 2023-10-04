VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Seq")
'@Exposed
'@Interface

' Public constants common to all Seq classes
' Implemented as anons.  This works because all the constants can be encapsulated by the long type
Public Enum SeqConst

    FIRST_INDEX = 1
    WHOLE_SEQ = -1
    TO_LAST_INDEX = 0
    DEFAULT_STEP = 1
    LACKS_ITEMS = -1
    INDEX_OUT_OF_BOUNDS = -1
    NOT_FOUND = -1

End Enum

Option Explicit

Public Function Deb(ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function Add(ByVal ipItem As Variant) As Long
End Function

Public Function Count(Optional ByRef ipItem As Variant) As Long
End Function

Public Function InsertAt(ByVal ipIndex As Long, ByVal ipItem As Variant) As Long
End Function

Public Property Get Item(ByVal ipIndex As Long) As Variant
End Property

Public Property Let Item(ByVal ipIndex As Long, ByVal ipItem As Variant)
End Property

Public Property Set Item(ByVal ipIndex As Long, ByVal ipItem As Variant)
End Property

Public Function Remove(ByRef ipItem As Variant) As Long
End Function

Public Function RemoveAt(ByRef ipIndex As Long) As Long
End Function

'Public Function Bind(ByRef ipArray As Variant) As Variant
'End Function

Public Function Clone() As Variant
End Function

Public Function Items() As Variant
End Function

Public Function Join(Optional ByRef ipSeparator As String = Chars.twComma) As String
End Function

Public Function Slice(Optional ByRef ipStart As Long = FIRST_INDEX, Optional ByRef ipRun As Long = WHOLE_SEQ, Optional ByRef ipStep As Long = DEFAULT_STEP) As Variant
End Function

Public Function ToArray() As Variant
End Function

Public Function ToCollection() As Collection
End Function

Public Function FirstIndex() As Long
End Function

Public Function LastIndex() As Long
End Function

Public Function FBOIndex() As Long
End Function

Public Function LBOIndex() As Long
End Function

Public Property Get First() As Variant
End Property

Public Property Let First(ByVal ipItem As Variant)
End Property

Public Property Set First(ByVal ipItem As Variant)
End Property

Public Property Get Last() As Variant
End Property

Public Property Let Last(ByVal ipItem As Variant)
End Property

Public Property Set Last(ByVal ipItem As Variant)
End Property

Public Property Get FBO() As Variant
End Property

Public Property Let FBO(ByVal ipItem As Variant)
End Property

Public Property Set FBO(ByVal ipItem As Variant)
End Property

Public Property Get LBO() As Variant
End Property

Public Property Let LBO(ByVal ipItem As Variant)
End Property

Public Property Set LBO(ByVal ipItem As Variant)
End Property

Public Function IndexOf(ByVal ipItem As Variant, Optional ByVal ipStart As Long = FIRST_INDEX, Optional ByVal ipRun As Long = TO_LAST_INDEX) As Long
End Function

Public Function LastIndexOf(ByVal ipItem As Variant, Optional ByVal ipStart As Long = FIRST_INDEX, Optional ByVal ipRun As Long = TO_LAST_INDEX) As Long
End Function

Public Function AddItems(ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function AddRange(ByVal ipRange As Variant) As Variant
End Function

Public Function InsertAtItems(ByVal ipIndex As Long, ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function InsertAtRange(ByVal ipIndex As Long, ByRef ipRange As Variant) As Variant
End Function

Public Function RemoveItems(ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function RemoveRange(ByRef ipRange As Variant) As Variant
End Function

Public Function RemoveAll(ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function Clear() As Variant
End Function

Public Function Reset() As Variant
End Function

Public Function RemoveIndexes(ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function RemoveIndexesRange(ByRef ipRange As Variant) As Variant
End Function

Public Function Fill(ByRef ipItem As Variant, Optional ByVal ipCount As Long = 1) As Variant
End Function

Public Function Head(Optional ByVal ipSplitAfter As Long = FIRST_INDEX) As Variant
End Function

Public Function Tail(Optional ByVal ipSplitAfter As Long = FIRST_INDEX) As Variant
End Function

Public Function RotLeft(Optional ByVal ipCount As Long = 1) As Variant
End Function

Public Function RotRight(Optional ByVal ipCount As Long = 1) As Variant
End Function

Public Function HoldsItem(ByRef ipItem As Variant) As Boolean
End Function

Public Function Contains(ByRef ipItem As Variant) As Boolean
End Function

Public Function LacksItem(ByRef ipItem As Variant) As Boolean
End Function

Public Function HoldsItems() As Boolean
End Function

Public Function LacksItems() As Boolean
End Function

Public Property Get TypeName() As String
End Property

Public Function MappedIt(ByRef ipMapper As IMapper) As Variant
End Function

Public Function MapIt(ByVal ipMapper As IMapper) As Variant
End Function

Public Function FilterIt(ByVal ipComparer As IComparer) As Variant
End Function

Public Function CountIt(ByVal ipComparer As IComparer) As Long
End Function

Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant
End Function

Public Function Push(ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function PushRange(ByRef ipRange As Variant) As Variant
End Function

Public Function Pop() As Variant
End Function

Public Function PopRange(Optional ByVal ipCount As Long = 1) As Variant
End Function

Public Function enQueue(ParamArray ipParamarray() As Variant) As Variant
End Function

Public Function EnqueueRange(ByRef ipRange As Variant) As Variant
End Function

Public Function Dequeue() As Variant
End Function

Public Function DequeueRange(Optional ByVal ipCount As Long = 1) As Variant
End Function

Public Function Sort() As Variant
End Function

Public Function Sorted() As Variant
End Function

Public Sub Swap(ByRef ipLHSIndex As Long, ByRef ipRHSIndex As Long)
End Sub

Public Function Reverse() As Variant
End Function

Public Function Reversed() As Variant
End Function

Public Function IsUnique() As Boolean
End Function

Public Function IsNotUnique() As Boolean
End Function

Public Function Dedup() As Variant
End Function

Public Function Deduped() As Variant
End Function

Public Function Freq() As KvpA
End Function

Public Function SetOf(ByRef ipSet As e_SetoF, ByRef ipRangeItem As ISeq) As ISeq
End Function
