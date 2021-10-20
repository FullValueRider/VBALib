VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableArrayToRowLysts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'@predeclaredId
Option Explicit
Option Private Module

Public Function ExecAction(ByVal ipTableArray As Variant) As Lyst
    
    Guard NotTableArray, Arrays.IsNotTableArray(ipTableArray), "VBALib.TableArrayToRowLysts.ExecAction", TypeName(Me)
    
    Dim myFirstRow As Long
    myFirstRow = LBound(ipTableArray, 1)
    
    Dim myLastRow As Long
    myLastRow = UBound(ipTableArray, 1)
    
    Dim myFirstColumn As Long
    myFirstColumn = LBound(ipTableArray, 2)
    
    Dim myLastColumn As Long
    myLastColumn = UBound(ipTableArray, 2)
    
    Dim myLystOfLyst As Lyst
    Set myLystOfLyst = Lyst.Deb
    Dim myRow As Long
    Dim myColumn As Long
    For myRow = myFirstRow To myLastRow
    
        Dim myRankLyst As Lyst
        Set myRankLyst = Lyst.Deb
        'myLystOfLyst.Add Lyst.Deb
        
        For myColumn = myFirstColumn To myLastColumn
            
            myRankLyst.AddValidatedItem ipTableArray(myRow, myColumn)
            'myLystOfLyst.Item(myRow - myFirstRow).Add ipTableArray(myRow, myColumn)
            
        Next
    
        myLystOfLyst.AddValidatedItem myRankLyst.Clone
        
    Next
    
    Set ExecAction = myLystOfLyst
    
End Function