VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EnumTableToLystActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
    
    '@Folder("VBALib.Enums")
    '@PredeclaredId
    '@Exposed
    Option Explicit
    
    Public Enum TableToLystAction
		
        RankIsColumnFirstItemIsNoAction
        RankIsRowFirstItemActionIsNoAction
		RankIsColumnFirstItemActionIsCopy
		RankIsRowFirstItemActionIsCopy
		RankIsColumnFirstItemActionIsSplit
		RankIsRowFirstItemActionIsSplit

	End Enum
    
    Private Const mcLocation      As String = "VBALib.EnumTableToLystActions."
    
    Private Type state
        
        Action                  As Scripting.Dictionary
        Strings                 As Scripting.dictionary
        
    End Type
    
    
    Private s                   As state
    
    
    Private Sub Class_Initialize()
        
        Guard InvalidNew, Not (Me Is Enums.TableToLystActions), mcLocation & "Class_Initialize"
        SetupDictionaries
        
    End Sub
    
    
    Private Sub SetupDictionaries()
        
        Set s.Action = CreateObject("Scripting.DIctionary")
        With s.Action
        
            .Add RankIsColumnFirstItemIsNoAction, New TableArrayToColumnLysts
            .Add RankIsRowFirstItemActionIsNoAction, New TableArrayToRowLysts
            .Add RankIsColumnFirstItemActionIsCopy, New TableArrayToColumnLystsCopyFirst
		    .Add RankIsRowFirstItemActionIsCopy, New TableArrayToRowLystsCopyFirst
		    .Add RankIsColumnFirstItemActionIsSplit, New TableArrayToColumnLystsSplitFirst
		    .Add RankIsRowFirstItemActionIsSplit, New TableArrayToRowLystsSplitFirst
        
        End With
        
        Set s.Strings = CreateObject("Scripting.Dictionary")
        With s.Strings
        
            .Add RankIsColumnFirstItemIsNoAction, "RankIsColumnNoFirstItemAction"
            .Add RankIsRowFirstItemActionIsNoAction, "ByRowIgnoreFist"
            .Add RankIsColumnFirstItemActionIsCopy, "RankIsColumnCopyFirstItem"
            .Add RankIsRowFirstItemActionIsCopy, "RankIsRowCopyFirstItem"
            .Add RankIsColumnFirstItemActionIsSplit, "RankIsColumnSplitFirstItem"
            .Add RankIsRowFirstItemActionIsSplit, "RankIsRowSplitFirstItem"
        
        End With
        
    End Sub
    
    Public Function Strings() As Variant
        Strings = s.Strings.Items
    End Function
    
    
    Public Function Members() As Variant
        Members = s.Action.Keys
    End Function
    
    
    Public Function Actions() As Variant
        Actions = s.Action.Items
    End Function
    
    
    Public Function Ord(ByVal ipEnum As TableToLystAction) As Long
        
        Dim myItem As Variant
        Dim myOrd As Long
        myOrd = 1
        For Each myItem In s.Strings.Keys
            
            If myItem = ipEnum Then
                
                Ord = myOrd
                Exit Function
                
            Else
            
                myOrd = myOrd + 1
                
            End If
            
        Next
        
    End Function
    
    Public Property Get AsEnum(ByVal ipEnum As TableToLystAction) As TableToLystAction
        AsEnum = ipEnum
    End Property
    
    
    Public Function CanExec(ByVal ipEnum As TableToLystAction) As Boolean
        CanExec = VBA.IsObject(s.Action.Item(ipEnum))
     End Function
    
    
    Public Property Get ToString(ByVal ipEnum As TableToLystAction) As String
        ToString = s.Strings.Item(ipEnum)
    End Property
        
    
    Public Function Exec(ByVal ipEnum As TableToLystAction, ByVal ipTableArray As Variant) As Lyst
        
        Guard ActionNotExecutable, Not CanExec(ipEnum), mcLocation & "Exec"
        ' Debug.Print TypeName(s.Action.Item(ipEnum))
        ' Dim myItems As Variant
        ' myItems = s.Action.Items
        ' Dim myObject As Object
        ' Set myObject = s.Action.Item(ipEnum)
        Set Exec = s.Action.Item(ipEnum).execaction(ipTableArray)
        
    End Function
    
    
    Public Property Get HoldsMember(ByVal ipValue As Long) As Boolean
        HoldsMember = s.Strings.Exists(ipValue)
    End Property
    
    Public Property Get LacksMember(ByVal ipValue As Long) As Boolean
        LacksMember = Not s.Strings.Exists(ipValue)
    End Property
    
    
    Public Property Get Name() As String
        Name = "TableToLystActions"
    End Property
    
    
    Public Property Get Count() As Long
        Count = s.Strings.Count
    End Property
    
    
    Public Function FirstItems() As Long
        FirstItems = 0
    End Function
    
    Public Function RankLysts() As Long
        RankLysts = 1
    End Function
    
End Class