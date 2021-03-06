Attribute VB_Name = "Comparer"
Option Explicit
'@Folder("VBALib")
Public Enum Action
        
        Equal
        NotEqual
        LessThan
        LessThanOrEqual
        NotMoreThan
        MoreThan
        MoreThanOrEqual
        NotLessThan
        
End Enum

Public Function Compare(ByVal ipComparer As Action, ByVal ipRef As Variant, ByVal ipTest As Variant) As Boolean
        
        Guard ResultCode.NotSameType, TypeName(ipRef) <> TypeName(ipTest), "VBALib.Comparer.Compare", Array(TypeName(ipRef), TypeName(ipTest))
        
        Dim myresult As Boolean
        
        Select Case ipComparer
                
                Case Action.Equal
                        
                        myresult = ipRef = ipTest
                
                        
                Case Action.NotEqual
                
                        myresult = ipRef <> ipTest
                        
                        
                Case Action.LessThan
                
                        myresult = ipRef < ipTest
                        
                        
                Case LessThanOrEqual, NotMoreThan
                
                        myresult = ipRef <= ipTest
                
                        
                Case MoreThan
                
                        myresult = ipRef > ipTest
                        
                        
                Case MoreThanOrEqual, NotLessThan
                
                        myresult = ipRef >= ipTest
                
                
        End Select
        
        Compare = myresult
        
End Function

        

