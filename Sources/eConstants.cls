VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ePresets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const ePresetName As String = "eConstants"

Public Enum PresetId
    
    NoValue = -1
    'Number related
    NaN
    MaxLong
    MinLong
    
    ' Array related
     ArrayFirstRank
    
    ' Result related
    ResultStatusOkay
    ResultStatusNotOkay

    'Kvp related
    KvpInsertIndexIsAdd
    DefaultDec
    DefaultInc
    DefaultAdjust
    
End Enum

Private Type State
    
    Values                  As Scripting.Dictionary
    Strings                 As Scripting.Dictionary
    
End Type

Private s                   As State

Private Sub Class_Initialize()
        
    'If BailOut.When(Not (Me Is enums.TableToLystActions), Msg.InvalidNew) Then Exit Sub
    SetupKvps
        
End Sub

Private Sub SetupKvps()
  
    Set s.Values = New Scripting.Dictionary
    With s.Values
        
        'Number related
        .Add Key:=NaN, Item:="NaN"
        .Add Key:=MaxLong, Item:=&H7FFFFFFF
        .Add Key:=MinLong, Item:=&HFFFFFFFF
        
        ' Array related
        .Add Key:=ArrayFirstRank, Item:=1
        
        ' Result related
        .Add Key:=ResultStatusOkay, Item:=True
        .Add Key:=ResultStatusNotOkay, Item:=False

        'Kvp related
        .Add Key:=KvpInsertIndexIsAdd, Item:=0
        .Add Key:=DefaultDec, Item:=1
        .Add Key:=DefaultInc, Item:=1
        .Add Key:=DefaultAdjust, Item:=1
        
    End With
    
    Set s.Strings = New Scripting.Dictionary
    With s.Strings
    ' Number related
        .Add Key:=NaN, Item:="NaN"
        .Add Key:=MaxLong, Item:="MaxLOng"
        .Add Key:=MinLong, Item:="MinLOng"
        
        ' Attay related
       .Add Key:=ArrayFirstRank, Item:="ArrayFirstRank"
        
        ' Result related
        .Add Key:=ResultStatusOkay, Item:="ResultStatusOkay"
        .Add Key:=ResultStatusNotOkay, Item:="ResultStatusNotOkay"

        'Kvp related
        .Add Key:=KvpInsertIndexIsAdd, Item:="KvpInsertIndexIsAdd"
        .Add Key:=DefaultDec, Item:="DefaultDec"
        .Add Key:=DefaultInc, Item:="DefaultInc"
        .Add Key:=DefaultAdjust, Item:="DefaultAdjust"
        
    End With
            
        
End Sub

Public Function Values() As Lyst
    Set Values = s.Values.ItemsRef.Clone
End Function

Public Function Members() As Lyst
    Set Members = s.Values.KeysRef.Clone
End Function


Public Function Ord(ByVal ipEnum As TableToLystAction) As Long
        
    Dim myItem As Variant
    Dim myOrd As Long
    myOrd = 1
    Ord = -1
    For Each myItem In s.Values.Keys
            
        If myItem = ipEnum Then
                
            Ord = myOrd
            Exit Function
                
                
        Else
            
            myOrd = myOrd + 1
                
                
        End If
            
    Next
        
End Function

Public Property Get AsEnum(ByVal ipEnum As PresetId) As PresetId
    AsEnum = ipEnum
End Property

Public Property Get Value(ByVal ipEnum As PresetId) As Variant
    Value = s.Values.Item(ipEnum)
End Property

Public Property Get ToString(ByVal ipEnum As TableToLystAction) As String
    ToString = s.Strings.Item(ipEnum)
End Property


Public Property Get HoldsMember(ByVal ipValue As Long) As Boolean
    HoldsMember = s.Strings.HoldsItem(ipValue)
End Property

Public Property Get LacksMember(ByVal ipValue As Long) As Boolean
    LacksMember = Not s.Strings.LacksItem(ipValue)
End Property

Public Property Get Name() As String
    Name = ePresetName
End Property

Public Property Get Count() As Long
    Count = s.Strings.Count
End Property
