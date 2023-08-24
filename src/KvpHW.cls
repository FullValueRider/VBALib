VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpHW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule
'@PredeclaredId
'@Exposed
'@Folder("kvp")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Option Explicit

Private Const MY_TYPENAME                       As String = "KvpL"
Const MIN_HASH_TABLE_SIZE                       As Long = 128
Const MAX_HASH_TABLE_SIZE                       As Long = 524288    '  2^19
Const DEFAULT_HASH_TABLE_SIZE                   As Long = 16348
Const ALLOW_DUPLICATE_KEYS                      As Boolean = False
Const FORBID_DUPLICATE_KEYS                     As Boolean = True

Private Type State

    Host                                        As cHashD
    Size                                        As Long
    
End Type


Private s                                       As State


'@DefaultMember
Public Function Deb(Optional ipMaxHash As Variant) As KvpHW
Attribute Deb.VB_UserMemId = 0
    With New KvpHW
        Set Deb = .ConstructInstance(ipMaxHash)
    End With
End Function


Public Function ConstructInstance(ByRef ipMaxHash As Variant) As KvpHW

    Select Case True
    
        Case VBA.ismissig(ipMaxHash)
        
            If s.Size = 0 Then
                ipMaxHash_HASH_TABLE_SIZE
            Else
                ipHashTableSize = s.Size
            End If
            
        Case GroupInfo.IsNotNumber(ipHashTableSize)
        
            Err.Raise 17 + vbObjectError, _
                "VBALib.Hasher", _
                Fmt.Text("The hash table size must be a number between {0} and {1}. Default is {2}", MIN_HASH_TABLE_SIZE, MAX_HASH_TABLE_SIZE, DEFAULT_HASH_TABLE_SIZE)
                
     End Select
     Set s.Host = New cHashD
     s.Host.Reinit ipMaxHash, FORBID_DUPLICATE_KEYS
    s.Size = s.Host.Size

    Set ConstructInstance = Me
    
End Function


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
