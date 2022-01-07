VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultParsedVariant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
''@IgnoreModule
'@Exposed
'@PredeclaredId
'@Folder("VBALib.Result")

Private Enum ItemsIndex
    
    StartIndex = 0
    RunLength
    
End Enum

Private Type Properties
    
    Items As ArrayList
    
    
End Type

Private p                           As Properties

Private Sub Class_Initialize()
    Set p.Items = New ArrayList
End Sub

Public Function Deb(ByVal ipitems As ArrayList) As ResultParsedVariant
    
    With New ResultParsedVariant
    
        Set Deb = .ReadyToUseInstance(ipitems)
    
    End With
    
End Function

Public Function ReadyToUseInstance(ByVal ipitems As ArrayList) As ResultParsedVariant
    
    Set p.Items = ipitems
    Set ReadyToUseInstance = Me
    
End Function


Public Function GetItems() As Lyst

    If Types.IsTypeByName(Types.Name.OfLyst, p.Items.Item(0)) Then
    
        Set GetItems = p.Items.Item(0)
        
    ElseIf Types.Group.IsNotIterable(p.Items.Item(0)) Then
    
        Set GetItems = Lyst.Deb.AddKnownItemsAsItems(p.Items.Item(0))
        
    Else
    
        Set GetItems = Lyst.Deb.AddKnownIterableAsItems(p.Items.Item(0))
        
    End If
    
End Function


Public Function GetKeys() As Lyst

   If Types.IsTypeByName(Types.Name.OfLyst, p.Items.Item(1)) Then
    
        Set GetKeys = p.Items.Item(1)
        
    ElseIf Types.Group.IsNotIterable(p.Items.Item(1)) Then
    
        Set GetKeys = Lyst.Deb.AddKnownItemsAsItems(p.Items.Item(1))
        
    Else
    
        Set GetKeys = Lyst.Deb.AddKnownIterableAsItems(p.Items.Item(1))
        
    End If
    
End Function

Public Property Get ResultItems() As Lyst
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set myLyst.Host = p.Items.Clone
    Set ResultItems = myLyst
End Property

Public Property Let ResultItems(ByVal ipLyst As Lyst)
    Set p.Items = ipLyst.Host.Clone
End Property
