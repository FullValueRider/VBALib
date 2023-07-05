VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
 Public Function IsAllocated(ByRef ipArray As Variant) As Boolean
Attribute IsAllocated.VB_Description = "Tests that an has been dimmed with valid bounds"

     If Not VBA.IsArray(ipArray) Then
        IsAllocated = False
    End If
    
     Dim myRanks As Long
     myRanks = Ranks(ipArray)
     IsAllocated = myRanks > 0
   
 End Function

 Public Function IsNotAllocated(ByRef ipArray As Variant) As Boolean
     IsNotAllocated = Not IsAllocated(ipArray)
 End Function
 
 
 Public Function IsNotQueryable(ByRef ipArray As Variant) As Boolean
    
    If Not VBA.IsArray(ipArray) Then
        IsNotQueryable = True
    End If
    
    IsNotQueryable = IsNotAllocated(ipArray)
     
 End Function
 
 
Public Function IsArray(ByRef ipArray As Variant, Optional ByRef ipArrayType As e_ArrayType = e_ArrayType.m_AnyArrayType) As Boolean
     
     If IsNotQueryable(ipArray) Then
        IsArray = False
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
     
     If IsNotAllocated(ipArray) Then
         IsNotArray = True
     End If
     
     Dim myRanks As Long
     myRanks = Ranks(ipArray)
     
    Select Case ipArrayType
        
        Case e_ArrayType.m_AnyArrayType
            myRanks = myRanks < 1
            
        Case e_ArrayType.m_ListArray
            myRanks = myRanks <> e_ArrayRanks.m_ListArrayRanks
            
        Case e_ArrayType.m_TableArray
            myRanks = myRanks <> e_ArrayRanks.m_TableArrayRanks
            
        Case e_ArrayType.m_MDArray
            myRanks = myRanks <= e_ArrayRanks.m_TableArrayRanks
            
        
    End Select
         
End Function
 
 
Public Function Count(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 0) As Long
     
    If IsNotQueryable(ipArray) Then
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
Public Function HasRank(ByRef ipArray As Variant, ByVal ipRank As Long) As Boolean
    
    Dim myRanks As Long
    myRanks = Ranks(ipArray)
    If myRanks < 1 Then
       HasRank = False
    End If
    
    If ipRank > myRanks Then
        HasRank = False
    End If
    
    HasRank = True
     
End Function

Public Function LacksRank(ByRef ipArray As Variant, ByVal ipRank As Long) As Boolean
    LacksRank = Not HasRank(ipArray, ipRank)
End Function


Public Function HasItems(ByRef ipArray As Variant) As Boolean
    
   If IsNotQueryable(ipArray) Then
       HasItems = False
   End If
   
   HasItems = Count(ipArray) > 1

End Function

' '@Description("True if ipArray is an array holding One item.  Optionally can query the GroupType of the single item")
' Public Function HasOneItem(ByRef ipArray As Variant) As Boolean
'
'     If IsNotQueryable(ipArray) Then Return False
'     Return Me.Count(ipArray) = 1
'
' End Function

' Public Function HasAnyItems(ByRef ipArray As Variant) As Boolean
'
'     If IsNotQueryable(ipArray) Then Return False
'     Return Count(ipArray) > 0
'
' End Function


' '@Description(True if the count Test and Ref is the same (Arrays can be compared with Ranks))
' Public Function IsSameCount _
' ( _
'     ByVal ipTestArray As Variant, _
'     ByVal ipRefArray As Variant, _
'     Optional ByVal ipTestRank As Long = 0, _
'     Optional ByVal ipRefRank As Long = 0 _
' ) As Boolean
'
'     If IsNotQueryable(ipTestArray) Then Return False
'     If IsNotQueryable(ipRefArray) Then Return False
'
'     Dim myRefCount As Long
'     If ipRefRank = 0 Then
'         myRefCount = Count(ipRefArray)
'
'     Else
'
'         Dim myRefRanks As Long = Ranks(ipRefArray)
'         If ipRefRank > myRefRanks Then
'              Return False
'         Else
'             myRefCount = Count(ipRefArray, ipRefRank)
'         End If
'
'     End If
'
'     Dim myTestCount As Long
'     If ipTestRank = 0 Then
'         myTestCount = Count(ipTestArray)
'
'     Else
'
'         Dim myTestRanks As Long = Ranks(ipTestArray)
'         If ipTestRank > myTestRanks Then
'              Return False
'         Else
'             myTestCount = Count(ipTestArray, ipTestRank)
'         End If
'
'     End If
'
'     Return myRefCount = myTestCount
'
' End Function


Public Function FirstIndex(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 1) As Variant
    
   If IsNotQueryable(ipArray) Then
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
    
   If IsNotQueryable(ipArray) Then
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


'@Ignore ParameterCanBeByVal
Public Function IsListArray(ByRef ipArray As Variant) As Boolean
    IsListArray = IsArray(ipArray, e_ArrayType.m_ListArray)
End Function

Public Function IsNotListArray(ByRef ipArray As Variant) As Boolean
    IsNotListArray = Not IsArray(ipArray, e_ArrayType.m_ListArray)
End Function

'@Ignore ParameterCanBeByVal
Public Function IsTableArray(ByRef ipArray As Variant) As Boolean
    IsTableArray = IsArray(ipArray, e_ArrayType.m_TableArray)
End Function

Public Function IsNotTableArray(ByRef ipArray As Variant) As Boolean
    IsNotTableArray = Not IsArray(ipArray, e_ArrayType.m_TableArray)
End Function

'@Ignore ParameterCanBeByVal
Public Function IsMDArray(ByRef ipArray As Variant) As Boolean
    IsMDArray = IsArray(ipArray, e_ArrayType.m_MDArray)
End Function

Public Function IsNotMDArray(ByRef ipArray As Variant) As Boolean
    IsNotMDArray = Not IsArray(ipArray, e_ArrayType.m_MDArray)
End Function
