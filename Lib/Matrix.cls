VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed

Private Const maColumnDimension     As Long = 2
Private Const maRowDimension        As Long = 1

Public Enum FirstItem

    IsSplit
    IsCopy
    
End Enum

Private Type State

    Table                           As Variant
    FirstRow                        As Long
    LastRow                         As Long
    FirstColumn                     As Long
    LastColumn                      As Long
    Transpose                       As Boolean
    
End Type

Private s                           As State


Public Function Deb(ByVal iparray As Variant) As Matrix

    With New Matrix
    
        Set Deb = .ReadyToUseInstance(iparray)
        
    End With
    
End Function

Public Function ReadyToUseInstance(ByVal iparray As Variant) As Matrix

    Set ReadyToUseInstance = Nothing
    
    If Types.Arrays.IsNotTableArray(iparray) Then Exit Function

    s.Table = iparray
    s.FirstRow = LBound(s.Table, maRowDimension)
    s.LastRow = UBound(s.Table, maRowDimension)
    s.FirstColumn = LBound(s.Table, maColumnDimension)
    s.LastColumn = UBound(s.Table, maColumnDimension)
    s.Transpose = False
    
    Set ReadyToUseInstance = Me
    
End Function


Public Sub Transpose()

    s.Transpose = Not s.Transpose
    
End Sub

'@DefaultMember
Public Property Get Item(ByVal iprow As Long, ByVal ipcol As Long) As Variant
Attribute Item.VB_UserMemId = 0

    Item = Empty
    If Me.LacksRow(iprow) Then Exit Property
    If Me.LacksColumn(ipcol) Then Exit Property
    
    If VBA.IsObject(s.Table(iprow, ipcol)) Then
    
        Set Item = s.Table(iprow, ipcol)
        
    Else
    
        Item = s.Table(iprow, ipcol)
        
    End If
        
End Property

Public Property Let Item(ByVal iprow As Long, ByVal ipcol As Long, ByVal ipValue As Variant)

    If Me.LacksRow(iprow) Then Exit Property
    If Me.LacksColumn(ipcol) Then Exit Property
    If VBA.IsObject(ipValue) Then
    
        Set s.Table(iprow, ipcol) = ipValue
        
    Else
    
        s.Table(iprow, ipcol) = ipValue
        
    End If

End Property


'Public Property Set Item(ByVal ipRow As Long, ByVal ipCol As Long, ByVal ipValue As Variant)
'
'    If Me.LacksRow(ipRow) Then Exit Property
'    If Me.LacksColumn(ipCol) Then Exit Property
'    Set s.Table(ipRow, ipCol) = ipValue
'
'End Property

Public Function Row(ByVal iprow As Long) As Lyst

    Set Row = Nothing
    If Me.LacksRow(iprow) Then Exit Function
    
    Dim myList As VBALib.Lyst
    Set myList = VBALib.Lyst.Deb
    
    Dim myCol As Long
    For myCol = s.FirstColumn To s.LastColumn
    
    ' Without myVar the code crashes at myList.add s.Table(myCol, ipRow)
        Dim myVar As Variant
        myVar = s.Table(iprow, myCol)
        myList.Add myVar
    
    Next
    
    Set Row = myList
    
End Function


Public Function Column(ByVal ipColumn As Long) As VBALib.Lyst

    Set Column = Nothing
    If Me.LacksColumn(ipColumn) Then Exit Function
    
    Dim myList As VBALib.Lyst
    Set myList = VBALib.Lyst.Deb
    
    Dim myRow As Long
    For myRow = s.FirstRow To s.LastRow
    
        Dim myVar As Variant
        myVar = s.Table(myRow, ipColumn)
        myList.Add myVar
    
    Next
    
    Set Column = myList
    
End Function


Public Function Rows() As VBALib.Lyst

    Dim myList As VBALib.Lyst
    Set myList = VBALib.Lyst.Deb
    
    Dim myRow As Long
    For myRow = s.FirstRow To s.LastRow
    
        myList.Add Me.Row(myRow)
        
    Next
    
    Set Rows = myList
    
End Function


Public Function Columns() As VBALib.Lyst

    Dim myList As VBALib.Lyst
    Set myList = VBALib.Lyst.Deb
    
    Dim myCol As Long
    For myCol = s.FirstColumn To s.LastColumn
    
        myList.Add Me.Column(myCol)
        
    Next
    
    Set Columns = myList
    
End Function


Public Function HasRow(ByVal iprow As Long) As Boolean
    HasRow = (iprow >= s.FirstRow) And (iprow <= s.LastRow)
End Function

Public Function LacksRow(ByVal iprow As Long) As Boolean
    LacksRow = Not HasRow(iprow)
End Function


Public Function HasColumn(ByVal ipColumn As Long) As Boolean
    HasColumn = (ipColumn >= s.FirstColumn) And (ipColumn <= s.LastColumn)
End Function

Public Function LacksColumn(ByVal ipColumn As Long) As Boolean
    LacksColumn = Not HasColumn(ipColumn)
End Function


 Public Function FirstRow() As Long
    FirstRow = s.FirstRow
 End Function
 
 Public Function LastRow() As Long
    LastRow = s.LastRow
 End Function
    
    
Public Function FirstCol() As Long
    FirstCol = s.FirstColumn
End Function

Public Function LastCol() As Long
    LastCol = s.LastColumn
End Function


Public Function Alpha() As Variant
    Alpha = s.Table(s.FirstRow, s.FirstColumn)
End Function

Public Function Omega() As Variant
    Omega = s.Table(s.LastRow, s.LastColumn)
End Function

Public Function KeyValueRows(ByVal ipFirstItem As FirstItem) As Variant

    Dim myKeys As Lyst
    Set myKeys = Lyst.Deb
    
    Dim myValues As Lyst
    Set myValues = Lyst.Deb
    
    Dim myRows As Lyst
    Set myRows = Rows
    Dim myRow As Variant
    For Each myRow In myRows
    
        
        myKeys.Add myRow.Item(0)
        
        If ipFirstItem = IsSplit Then myRow.RemoveAt 0
            
        myValues.Add myRow
        
    Next
    
    KeyValueRows = Array(myKeys, myValues)
    
End Function


Public Function KeyValueColumns(ByVal ipFirstItem As FirstItem) As Variant

    Dim myKeys As Lyst
    Set myKeys = Lyst.Deb
    
    Dim myValues As Lyst
    Set myValues = Lyst.Deb
    
    Dim myCols As Lyst
    Set myCols = Columns
    Dim myCol As Variant
    For Each myCol In myCols
    
        
        myKeys.Add myCol.Item(0)
        
        If ipFirstItem = IsSplit Then myCol.RemoveAt 0
            
        myValues.Add myCol
        
    Next
    
    KeyValueColumns = Array(myKeys, myValues)
    
End Function
