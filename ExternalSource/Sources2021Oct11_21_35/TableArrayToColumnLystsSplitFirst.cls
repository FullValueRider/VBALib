VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableArrayToColumnLystsSplitFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
    
'@PredeclaredId
'@Exposed
Option Explicit

Public Function ExecAction(ByVal ipTablearray As Variant) As Lyst
    
    Dim myLystOfRankLysts As Lyst
    Set myLystOfRankLysts = TableArrayToColumnLysts.ExecAction(ipTablearray)
    
    Dim myRankFirstItemsLyst As Lyst
    Set myRankFirstItemsLyst = Lyst.Deb
    
    Dim myRankLysts As Lyst
    Set myRankLysts = Lyst.Deb
    
    Dim myItem As Variant
    
    For Each myItem In myLystOfRankLysts
        
        Dim myInnerLyst As Lyst
        Set myInnerLyst = myItem
        myRankFirstItemsLyst.AddValidatedItem myInnerLyst.First
        myRankLysts.AddValidatedItem myInnerLyst.GetRange(1, myInnerLyst.Count - 1)
        
    Next
    
    Set ExecAction = Lyst.Deb.Add(myRankFirstItemsLyst).Add(myRankLysts)
    
End Function