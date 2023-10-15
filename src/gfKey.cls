VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "gfKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Return the item at the specified index"
'@PredeclaredId
'@Exposed
'@ModuleDescription("Return the item at the specified index")
'@Folder("Functors")
Option Explicit

Implements IMapper
Implements IReducer

Const MY_TYPENAME                                   As String = "gfKey"
Private Type State
    Key                                             As Variant
End Type

Private s                                           As State
'@DefaultMember
Public Function Deb(ByRef ipKey As Variant) As gfKey
Attribute Deb.VB_UserMemId = 0
    With New gfKey
        Set Deb = .ConstructInstance(ipKey)
    End With
End Function


Friend Function ConstructInstance(ByRef ipKey As Variant) As gfKey

    If VBA.IsObject(ipKey) Then
        Set s.Key = ipKey
    Else
        s.Key = ipKey
    End If
    
    Set ConstructInstance = Me
    
End Function

Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property
    

Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property

Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
    IMapper_ExecMapper = Item(ipHostItem)
End Function


Private Function IReducer_ExecReduction(ByRef ipHostItem As Variant) As Variant
    IReducer_ExecReduction = Item(ipHostItem)
End Function

Public Function Item(ByRef ipHostItem As Variant) As Variant

    Dim myNullResult() As Variant: myNullResult = Array(Null)
    
    Select Case GroupInfo.Id(ipHostItem)
    
        Case e_Group.m_array
        
            If GroupInfo.IsNotNumber(s.Key) Then
                Item = myNullResult
                Exit Function
           End If
        
            Dim myRanks As Long: myRanks = ArrayOp.Ranks(ipHostItem)
            If myRanks <> 1 Then
                Err.Raise 17 + vbObjectError, _
                    "VBALib.gfKey.Item", _
                    "Expecting array with Ranks = 1. Got Ranks = {0}", myRanks
            End If
            
            Item = Array(ipHostItem(s.Key))
            
        Case e_Group.m_ItemByIndex
        
            If GroupInfo.IsNotNumber(s.Key) Then
                Item = myNullResult
                Exit Function
            End If
            
            If VBA.TypeName(ipHostItem) = "Collection" Then
                Err.Raise 17 + vbObjectError, _
                    "VBALib.gfKey.Item", _
                    Fmt.Text("IpHostItem was a collection object. {nl}The collection object does not support the FirstIndex/LastIndex methods.{nl}Consider using the Bind method of the SeqC class")
                    
            End If
            
            Guard.IndexOutOfBounds s.Key, ipHostItem, "gfKey.Item", ipreportback:=Guard.REPORT_BACK
            
            Item = Array(ipHostItem.Item(s.Key))
             
        Case e_Group.m_ItemByKey
        
            
            If ipHostItem.LacksKey(s.Key) Then
                
                If GroupInfo.IsNumber(s.Key) Then
                    Dim myKeys() As Variant: myKeys = ipHostItem.Items
                    Item = Array(myKeys(VBA.CLng(s.Key)))
                    Exit Function
                End If
                
                Item = myNullResult
                Exit Function
            End If
            
            Item = Array(ipHostItem.Item(s.Key))
        
            
        Case e_Group.m_String
        
            If GroupInfo.IsNotNumber(s.Key) Then
                Item = myNullResult
                Exit Function
            End If
            
            If VBA.Len(ipHostItem) = 0 Then
                Item = myNullResult
            End If
        
            Item = Array(VBA.Mid$(ipHostItem, s.Key, 1))
            
        Case Else
            Item = myNullResult
            Err.Raise 17 + vbObjectError, _
                "VBALib.gfKey.Item", _
                Fmt.Text("Expecting an indexable Type or string.  Got {0}", VBA.TypeName(ipHostItem))
    End Select

End Function
