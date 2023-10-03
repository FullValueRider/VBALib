VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeqOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Seq")

Option Explicit


'@Description("Transposes a seq of seq. Non existent items are replaced by Empty")
Public Function Transpose(ByRef ipSeq As Variant) As Variant
Attribute Transpose.VB_Description = "Transposes a seq of seq. Non existent items are replaced by Empty"
    
    Dim myS As Variant: Set myS = ipSeq.Clone
    'Fmt.Dbg "{0}", myS
    'find the maximum size of the inner sequences
    Dim myMaxRowIndex As Variant: myMaxRowIndex = myS.ReduceIt(rdInner(rdCountIt)).ReduceIt(rdFoldl(fnMaxNum, Globals.minLong))
    
    ' The Transpose function cannot assume that the inner sequences are of equal length
    ' make the inner sequences equal in length.
    Dim myIndex As Long
    For myIndex = ipSeq.FirstIndex To ipSeq.LastIndex
        Dim myCount As Long: myCount = ipSeq.Item(myIndex).LastIndex
        If myCount < myMaxRowIndex Then
            ipSeq.Item(myIndex).Fill Empty, myMaxRowIndex - myCount
        End If
    Next
    
    
    ' create the recipient seq
    Dim myTrans As Variant: Set myTrans = ipSeq.Deb
    For myIndex = 1 To myMaxRowIndex
        myTrans.Add ipSeq.First.Deb.Fill(Empty, ipSeq.LastIndex)
    Next
    
    ' now do the transposition
    Dim myRowIndex As Long
    For myRowIndex = ipSeq.FirstIndex To ipSeq.LastIndex
    
        Dim myRow As ISeq: Set myRow = ipSeq.Item(myRowIndex)
        
        Dim myColIndex As Long
        For myColIndex = myRow.FirstIndex To myMaxRowIndex
            
            If VBA.IsObject(ipSeq.Item(myRowIndex).Item(myColIndex)) Then
                Set myTrans.Item(myColIndex).Item(myRowIndex) = ipSeq.Item(myRowIndex).Item(myColIndex)
            Else
                 myTrans.Item(myColIndex).Item(myRowIndex) = ipSeq.Item(myRowIndex).Item(myColIndex)
            End If
 
        Next
        
    Next
    'Fmt.Dbg "{0}", myTrans
    Set ipSeq = myTrans
    Set Transpose = ipSeq

End Function
