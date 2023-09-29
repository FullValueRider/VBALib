VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Permutations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Helpers")
Option Explicit

Private Type State
    Perms                           As SeqC
    Arr                             As Variant
End Type

Private s                           As State


' Generating permutation using Heap Algorithm
'https://en.wikipedia.org/wiki/Heap%27s_algorithm
'
' On the first call k= ubound of the array
' where indexing starts at 0
'procedure generate(k : integer, A : array of any):
'    If k = 1 Then
'        output (A)
'    Else
'        // Generate permutations with k-th unaltered
'        // Initially k = length(A)
'        generate(k - 1, A)
'
'        // Generate permutations for k-th swapped with each k-1 initial
'        for i := 0; i < k-1; i += 1 do
'            // Swap choice dependent on parity of k (even or odd)
'            If k Is Even Then
'                swap(A[i], A[k-1]) // zero-indexed, the k-th is at k-1
'            Else
'                swap(A[0], A[k-1])
'            End If
'            generate(k - 1, A)
'        end for
'    End If
' the items to be permuted are provided in ipInput
' the return value is a seq of seq which contains the permuted
' indexes in ipSeq, ie. the return is a seq of seq of long
' this allows the function to permute any
 Private Sub ByIndex(ByVal ipSize As Long)
    
    Dim myTmp As Long
    If ipSize = 1 Then
        s.Perms.Add SeqC(s.Arr)
    Else
        Dim myLastIndex As Long: myLastIndex = ipSize - 1
        ByIndex myLastIndex
        Dim i As Long
        For i = 0 To myLastIndex - 1
            If Maths.IsEven(ipSize) Then
                myTmp = s.Arr(i)
                s.Arr(i) = s.Arr(myLastIndex)
                s.Arr(myLastIndex) = myTmp
            Else
                myTmp = s.Arr(0)
                s.Arr(0) = s.Arr(myLastIndex)
                s.Arr(myLastIndex) = myTmp
            End If
            
            ByIndex myLastIndex

        Next

    End If

End Sub

' Returns a SeqC of Arrays of Long where each Array contains a permutation of the range 0 to ipSize -1
' ByIndexFrom0 3 returns  [0,1,2],[1,0,2],[2,0,1],[0,2,1],[1,2,0],[2,1,0]
    
Public Function ByIndexFrom0(ByRef ipSize As Long) As SeqC
    
    ReDim s.Arr(0 To ipSize - 1)
    
    Dim myIndex As Long
    For myIndex = 0 To ipSize - 1
        s.Arr(myIndex) = myIndex
    Next
    
    Set s.Perms = SeqC.Deb
    
    ByIndex ipSize
    
    Set ByIndexFrom0 = s.Perms
    
End Function

' Returns a SeqC of Arrays of Long where each inner Array is a permutation of the range 1 to ipSize
' ByIndexFrom1 3 returns {[1,2,3],[2,1,3],[3,1,2],[1,3,2],[2,3,1],[3,2,1]}
Public Function ByIndexFrom1(ByRef ipSize As Long) As SeqC
    
    ReDim s.Arr(0 To ipSize - 1)
    
    Dim myIndex As Long
    For myIndex = 0 To ipSize - 1
        s.Arr(myIndex) = myIndex + 1
    Next
    
    Set s.Perms = SeqC.Deb
    
    ByIndex ipSize
    
    Set ByIndexFrom1 = s.Perms
    
End Function

Public Function ByArray(ByRef ipArray As Variant) As SeqC

    If ArrayOp.IsNotArray(ipArray) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Permutations.ByArray", _
            Fmt.Text("Expecting an array.  Got {0}", VBA.TypeName(ipArray))
    End If
    
    Dim myA As Variant: myA = ArrayOp.Linearise(ipArray, ipBase:=0)
    Dim mySize As Long: mySize = ArrayOp.Count(ipArray)
    
    Dim myS As SeqC: Set myS = ByIndexFrom0(mySize)
    
    Dim myArrays As IterItems: Set myArrays = IterItems(myS)
    Do
        Dim myElements As IterItems: Set myElements = IterItems(myArrays.CurItem(0))
        Do
            Dim myE() As Variant: myE = Array(myElements.CurItem(0))
            If VBA.IsObject(myA(myE(0))) Then
                Set myElements.CurItem(0) = myA(myE(0))
            Else
                myElements.CurItem(0) = myA(myE(0))
            End If
        Loop While myElements.MoveNext
    Loop While myArrays.MoveNext
    Debug.Print
    Set ByArray = myS
    
End Function

Public Function ByItem(ByRef ipListType As Variant) As SeqC

    If GroupInfo.IsNotList(ipListType) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Permutations.ByItem", _
            Fmt.Text("Expecting an List Type.  Got {0}", VBA.TypeName(ipListType))
    End If
    
    Dim myS As SeqC
    If GroupInfo.IsIndexFrom0(ipListType) Then
   
        Set myS = ByIndexFrom0(ipListType.Count - 1)
    Else
        Set myS = ByIndexFrom1(ipListType.Count)
    End If
    
    Dim mySeqs As IterItems: Set mySeqs = IterItems(myS)
    Do
        Dim myElements As IterItems: Set myElements = IterItems(mySeqs.CurItem(0))
        Do
            Dim myE() As Variant: myE = Array(myElements.CurItem(0))
            If VBA.IsObject(ipListType.Item(myE(0))) Then
                Set myElements.CurItem(0) = ipListType.Item(myE(0))
            Else
                myElements.CurItem(0) = ipListType.Item(myE(0))
            End If
        Loop While myElements.MoveNext
    Loop While mySeqs.MoveNext
    Debug.Print
    Set ByItem = myS
    
End Function

Public Function ByKey(ByRef ipDicType As Variant) As SeqC

    If GroupInfo.IsNotDictionary(ipDicType) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Permutations.ByItem", _
            Fmt.Text("Expecting an Dictionary Type.  Got {0}", VBA.TypeName(ipDicType))
    End If
    
    Dim myS As SeqC
    ' Keys are indexed from 1
    Set myS = ByIndexFrom1(ipDicType.Count)
    
    Dim myKeys() As Variant: myKeys = ipDicType.Keys
    
    Dim mySeqs As IterItems: Set mySeqs = IterItems(myS)
    Do
        Dim myElements As IterItems: Set myElements = IterItems(mySeqs.CurItem(0))
        Do
            Dim myE() As Variant: myE = Array(myElements.CurItem(0))
            If VBA.IsObject(myKeys(myE(0))) Then
                Set myElements.CurItem(0) = myKeys(myE(0))
            Else
                myElements.CurItem(0) = myKeys(myE(0))
            End If
        Loop While myElements.MoveNext
    Loop While mySeqs.MoveNext
    Debug.Print
    Set ByKey = myS
    
End Function
