VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Lib")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit
 'The 'Arrays' class provides a collection of functions that take an array and return information about the array.

 Public Enum e_ArrayType
     
     m_AnyArrayType = -1
     m_ListArray = 0                 ' One dimension array
     m_TableArray = 1                ' Two dimension array
     m_MDArray = 2                   ' More than two dimension array

 End Enum


 Public Enum e_ArrayRanks
     m_ListArrayRanks = 1
     m_TableArrayRanks = 2
     m_MDArrayRanks = 3
 End Enum

 
'@Description("Tests that an has been dimmed with valid bounds")
Public Function LacksItems(ByRef ipArray As Variant) As Boolean
Attribute LacksItems.VB_Description = "Tests that an has been dimmed with valid bounds"

    If Not VBA.IsArray(ipArray) Then
       LacksItems = True
    Else
        LacksItems = Ranks(ipArray) < 1
    End If
  
End Function

Public Function HoldsItems(ByRef ipArray As Variant) As Boolean
    HoldsItems = Not LacksItems(ipArray)
End Function
 
 
Public Function HoldsOneItem(ByRef ipArray As Variant) As Boolean
    HoldsOneItem = Count(ipArray) = 1
End Function

Public Function HoldsItem(ByRef ipArray As Variant, ByRef ipItem As Variant) As Boolean

    Dim myItem As Variant
    For Each myItem In ipArray
        If ComparerHelpers.Equals(myItem, ipItem) Then
            HoldsItem = True
            Exit Function
        End If
    Next
    
    HoldsItem = False
    
End Function

Public Function LacksItem(ByRef ipArray As Variant, ByRef ipItem As Variant) As Boolean
    LacksItem = Not HoldsItem(ipArray, ipItem)
End Function

Public Function IsArray(ByRef ipArray As Variant, Optional ByRef ipArrayType As e_ArrayType = e_ArrayType.m_AnyArrayType) As Boolean
     
     IsArray = False
     
     If Not VBA.IsArray(ipArray) Then
         Exit Function
    End If
    
    If LacksItems(ipArray) Then
        Exit Function
    End If
     
     Dim myRanks As Long
     myRanks = Ranks(ipArray)
     
     Select Case ipArrayType
         
         Case m_AnyArrayType:                   IsArray = True
         Case e_ArrayType.m_ListArray:          IsArray = myRanks = m_ListArrayRanks
         Case e_ArrayType.m_TableArray:         IsArray = myRanks = m_TableArrayRanks
         Case e_ArrayType.m_MDArray:            IsArray = myRanks > m_TableArrayRanks
         
         
     End Select
     
End Function


Public Function IsNotArray(ByRef ipArray As Variant, Optional ByRef ipArrayType As e_ArrayType = m_AnyArrayType) As Boolean
     IsNotArray = Not IsArray(ipArray, ipArrayType)
End Function
 
Public Function GetIndex(ByRef ipArray As Variant, ParamArray ipDims() As Variant) As Long
' ToDo: this is where conversion from multi dimension to linear dimension occurs.
' some dummy statements to pacify rubberduck
    Dim myDummy As Long
    ipDims(0) = ipArray
    myDummy = ipDims(0)
    GetIndex = myDummy
End Function

Public Function Count(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 0) As Long
     
    If LacksItems(ipArray) Then
        Count = -1
    End If
 
    Count = CountKnownArray(ipArray, ipRank)
     
End Function
     
 
Friend Function CountKnownArray(ByVal ipArray As Variant, Optional ByVal ipRank As Long = 0) As Long
     
     Dim mySize As Long
     If ipRank = 0 Then ' Count all elements of the array
         
         mySize = 1
         Dim myRank As Long
         Dim myRanks As Long
         myRanks = Me.Ranks(ipArray)
         
         For myRank = 1 To myRanks
             mySize = mySize * (UBound(ipArray, myRank) - LBound(ipArray, myRank) + 1)
         Next
         
     Else
         
         If Me.LacksRank(ipArray, ipRank) Then
             mySize = 0
         Else
             mySize = UBound(ipArray, ipRank) - LBound(ipArray, ipRank) + 1
         End If
         
     End If
             
    CountKnownArray = mySize
     
End Function

'Public Function HoldsOneItem(ByRef ipArray As Variant) As Boolean
'    If HoldsItems(ipArray) Then
'        HoldsOneItem = Count(ipArray) = 1
'    Else
'        HoldsOneItem = False
'    End If
'End Function

 '@Description("Returns the number of dimensions of an array. Returns -1 if not an array")
Public Function Ranks(ByVal ipArray As Variant) As Long
Attribute Ranks.VB_Description = "Returns the number of dimensions of an array. Returns -1 if not an array"

     If Not VBA.IsArray(ipArray) Then
         Ranks = -1
     End If
     
     Dim myIndex As Long
     For myIndex = 1 To 60000
     
         On Error Resume Next

         Dim myDummy As Long
         myDummy = UBound(ipArray, myIndex)
         
         If Err.Number <> 0 Or myDummy = -1 Then
             On Error GoTo 0
             Exit For
         End If
         
         Err.Clear
         
     Next
     
     On Error GoTo 0
     
     Ranks = myIndex - 1
     
End Function


 '@Ignore ParameterCanBeByVal
Public Function HoldsRank(ByRef ipArray As Variant, ByVal ipRank As Long) As Boolean
    
    Dim myRanks As Long
    myRanks = Ranks(ipArray)
    If myRanks < 1 Then
       HoldsRank = False
    End If
    
    If ipRank > myRanks Then
        HoldsRank = False
    End If
    
    HoldsRank = True
     
End Function

Public Function LacksRank(ByRef ipArray As Variant, ByVal ipRank As Long) As Boolean
    LacksRank = Not HoldsRank(ipArray, ipRank)
End Function



Public Function FirstIndex(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 1) As Variant
    
   If LacksItems(ipArray) Then
       FirstIndex = Null
   End If
   
   Dim myRanks As Long
   myRanks = Ranks(ipArray)
   
   If ipRank >= 1 And ipRank <= myRanks Then
       FirstIndex = LBound(ipArray, ipRank)
   Else
       FirstIndex = Null
   End If
    
End Function

Public Function LastIndex(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 1) As Variant
    
   If LacksItems(ipArray) Then
       LastIndex = Null
   End If
   
   Dim myRanks As Long
   myRanks = Ranks(ipArray)
   
   If ipRank >= 1 And ipRank <= myRanks Then
       LastIndex = UBound(ipArray, ipRank)
   Else
       LastIndex = Null
   End If
   
End Function



