VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ranges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed


'@Ignore AssignedByValParameter
Public Function GetSeries(ByVal ipstart As Variant, ByVal ipCount As Long, Optional ByVal ipStep As Variant = 1) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set GetSeries = myLyst
    
    If Types.Group.IsNotNumber(ipstart, ipCount, ipStep) Then Exit Function
    
    Dim myItem As Variant
    For myItem = 0 To ipCount - 1
    
        myLyst.Add ipstart
        ipstart = ipstart + ipStep
        
    Next
    
End Function


Public Function GetConstSeries(ByVal ipCount As Long, ByVal ipValue As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set GetConstSeries = myLyst
    
    Dim myItem As Long
    For myItem = 0 To ipCount - 1
    
        myLyst.Add ipValue
        
    Next
    
End Function
