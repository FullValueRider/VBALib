VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrExDialogOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *****************************************************************************
' *  VCOM library: vbWatchdog  http://www.everythingaccess.com/vbwatchdog.htm *
' *  3.9.2 (Ultimate Edition TRI built May 26 2021                            *
' *---------------------------------------------------------------------------*
' *  This is a demonstration version of the library for TESTING purposes only.*
' *                                                                           *
' *  This software is provided "as-is," without any express or implied        *
' *  warranty.  In no event shall the author be held liable for any           *
' *  consequential, incidental, direct, indirect, special, punitive, or other *
' *  damages arising from the use of this software (including without         *
' *  limitation, damages for loss of business profits, business interruption, *
' *  loss of information or other pecuniary loss).                            *
' *****************************************************************************

Option Explicit

Enum DefaultButtonActions
    BUTTONACTION_ONERRORRESUMENEXT = &H2
    BUTTONACTION_ONERROREND = &H4
    BUTTONACTION_ONERRORDEBUG = &H5
    BUTTONACTION_ONERRORRETRY = &H7
    BUTTONACTION_ONERROREXITPROCEDURE = &H9
    BUTTONACTION_SHOWVARIABLES = &HA
    BUTTONACTION_SHOWHELP = &HB
    BUTTONACTION_ONERRORUNWIND = &H11
    BUTTONACTION_ONERRORUNWINDNOFINALLY = &H12
    BUTTONACTION_VARIABLES_CLOSE = &H101
End Enum

Public VCOMObject As Object      ' DO NOT USE!!! THIS IS FOR INTERNAL USE ONLY!!!
Attribute VCOMObject.VB_VarUserMemId = -4

Property Get HTML_MainBody() As String
    HTML_MainBody = VCOMObject.HTML_MainBody
End Property

Property Let HTML_MainBody(ByVal Value As String)
    VCOMObject.HTML_MainBody = Value
End Property

Property Get HTML_MoreInfoBody() As String
    HTML_MoreInfoBody = VCOMObject.HTML_MoreInfoBody
End Property

Property Let HTML_MoreInfoBody(ByVal Value As String)
    VCOMObject.HTML_MoreInfoBody = Value
End Property

Property Get HTML_CallStackItem() As String
    HTML_CallStackItem = VCOMObject.HTML_CallStackItem
End Property

Property Let HTML_CallStackItem(ByVal Value As String)
    VCOMObject.HTML_CallStackItem = Value
End Property

Property Get HTML_VariableItem() As String
    HTML_VariableItem = VCOMObject.HTML_VariableItem
End Property

Property Let HTML_VariableItem(ByVal Value As String)
    VCOMObject.HTML_VariableItem = Value
End Property

Property Get WindowCaption() As String
    WindowCaption = VCOMObject.WindowCaption
End Property

Property Let WindowCaption(ByVal Value As String)
    VCOMObject.WindowCaption = Value
End Property

Property Get MinimumWindowWidth() As Long
    MinimumWindowWidth = VCOMObject.MinimumWindowWidth
End Property

Property Let MinimumWindowWidth(ByVal Value As Long)
    VCOMObject.MinimumWindowWidth = Value
End Property

Property Get MoreInfoCaption() As String
    MoreInfoCaption = VCOMObject.MoreInfoCaption
End Property

Property Let MoreInfoCaption(ByVal Value As String)
    VCOMObject.MoreInfoCaption = Value
End Property

Property Get LessInfoCaption() As String
    LessInfoCaption = VCOMObject.LessInfoCaption
End Property

Property Let LessInfoCaption(ByVal Value As String)
    VCOMObject.LessInfoCaption = Value
End Property

Property Get ButtonPaddingH() As Long
    ButtonPaddingH = VCOMObject.ButtonPaddingH
End Property

Property Let ButtonPaddingH(ByVal Value As Long)
    VCOMObject.ButtonPaddingH = Value
End Property

Property Get ButtonPaddingV() As Long
    ButtonPaddingV = VCOMObject.ButtonPaddingV
End Property

Property Let ButtonPaddingV(ByVal Value As Long)
    VCOMObject.ButtonPaddingV = Value
End Property

Property Get ButtonSpacingH() As Long
    ButtonSpacingH = VCOMObject.ButtonSpacingH
End Property

Property Let ButtonSpacingH(ByVal Value As Long)
    VCOMObject.ButtonSpacingH = Value
End Property

Property Get ButtonSpacingV() As Long
    ButtonSpacingV = VCOMObject.ButtonSpacingV
End Property

Property Let ButtonSpacingV(ByVal Value As Long)
    VCOMObject.ButtonSpacingV = Value
End Property

Property Get PaddingH() As Long
    PaddingH = VCOMObject.PaddingH
End Property

Property Let PaddingH(ByVal Value As Long)
    VCOMObject.PaddingH = Value
End Property

Property Get PaddingV() As Long
    PaddingV = VCOMObject.PaddingV
End Property

Property Let PaddingV(ByVal Value As Long)
    VCOMObject.PaddingV = Value
End Property

Property Get ScreenBorderPaddingV() As Long
    ScreenBorderPaddingV = VCOMObject.ScreenBorderPaddingV
End Property

Property Let ScreenBorderPaddingV(ByVal Value As Long)
    VCOMObject.ScreenBorderPaddingV = Value
End Property

Property Get ColumnPaddingH() As Long
    ColumnPaddingH = VCOMObject.ColumnPaddingH
End Property

Property Let ColumnPaddingH(ByVal Value As Long)
    VCOMObject.ColumnPaddingH = Value
End Property

Property Get LineSpacingV() As Long
    LineSpacingV = VCOMObject.LineSpacingV
End Property

Property Let LineSpacingV(ByVal Value As Long)
    VCOMObject.LineSpacingV = Value
End Property

Property Get MainBackColor() As Long
    MainBackColor = VCOMObject.MainBackColor
End Property

Property Let MainBackColor(ByVal Value As Long)
    VCOMObject.MainBackColor = Value
End Property

Property Get MainBackColor2() As Long
    MainBackColor2 = VCOMObject.MainBackColor2
End Property

Property Let MainBackColor2(ByVal Value As Long)
    VCOMObject.MainBackColor2 = Value
End Property

Property Get MainBackFillType() As Long
    MainBackFillType = VCOMObject.MainBackFillType
End Property

Property Let MainBackFillType(ByVal Value As Long)
    VCOMObject.MainBackFillType = Value
End Property

Property Get MoreInfoBackColor2() As Long
    MoreInfoBackColor2 = VCOMObject.MoreInfoBackColor2
End Property

Property Let MoreInfoBackColor2(ByVal Value As Long)
    VCOMObject.MoreInfoBackColor2 = Value
End Property

Property Get MoreInfoBackFillType() As Long
    MoreInfoBackFillType = VCOMObject.MoreInfoBackFillType
End Property

Property Let MoreInfoBackFillType(ByVal Value As Long)
    VCOMObject.MoreInfoBackFillType = Value
End Property

Property Get ButtonBarBackColor() As Long
    ButtonBarBackColor = VCOMObject.ButtonBarBackColor
End Property

Property Let ButtonBarBackColor(ByVal Value As Long)
    VCOMObject.ButtonBarBackColor = Value
End Property

Property Get ButtonBarBackColor2() As Long
    ButtonBarBackColor2 = VCOMObject.ButtonBarBackColor2
End Property

Property Let ButtonBarBackColor2(ByVal Value As Long)
    VCOMObject.ButtonBarBackColor2 = Value
End Property

Property Get ButtonBarBackFillType() As Long
    ButtonBarBackFillType = VCOMObject.ButtonBarBackFillType
End Property

Property Let ButtonBarBackFillType(ByVal Value As Long)
    VCOMObject.ButtonBarBackFillType = Value
End Property

Property Get MaxNumCallStackItems() As Long
    MaxNumCallStackItems = VCOMObject.MaxNumCallStackItems
End Property

Property Let MaxNumCallStackItems(ByVal Value As Long)
    VCOMObject.MaxNumCallStackItems = Value
End Property

Property Get MaxNumVariablesItems() As Long
    MaxNumVariablesItems = VCOMObject.MaxNumVariablesItems
End Property

Property Let MaxNumVariablesItems(ByVal Value As Long)
    VCOMObject.MaxNumVariablesItems = Value
End Property

Property Get DefaultButtonIsBold() As Boolean
    DefaultButtonIsBold = VCOMObject.DefaultButtonIsBold
End Property

Property Let DefaultButtonIsBold(ByVal Value As Boolean)
    VCOMObject.DefaultButtonIsBold = Value
End Property

Property Get ShowMoreInfoButton() As Boolean
    ShowMoreInfoButton = VCOMObject.ShowMoreInfoButton
End Property

Property Let ShowMoreInfoButton(ByVal Value As Boolean)
    VCOMObject.ShowMoreInfoButton = Value
End Property

Property Get AllowEnterKey() As Boolean
    AllowEnterKey = VCOMObject.AllowEnterKey
End Property

Property Let AllowEnterKey(ByVal Value As Boolean)
    VCOMObject.AllowEnterKey = Value
End Property

Property Get AllowTabKey() As Boolean
    AllowTabKey = VCOMObject.AllowTabKey
End Property

Property Let AllowTabKey(ByVal Value As Boolean)
    VCOMObject.AllowTabKey = Value
End Property

Property Get DefaultButtonID() As Long
    DefaultButtonID = VCOMObject.DefaultButtonID
End Property

Property Let DefaultButtonID(ByVal Value As Long)
    VCOMObject.DefaultButtonID = Value
End Property

Property Get CustomImageTransparentColor() As Long
    CustomImageTransparentColor = VCOMObject.CustomImageTransparentColor
End Property

Property Let CustomImageTransparentColor(ByVal Value As Long)
    VCOMObject.CustomImageTransparentColor = Value
End Property

Property Get OnKeyDown() As String
    OnKeyDown = VCOMObject.OnKeyDown
End Property

Property Let OnKeyDown(ByVal Value As String)
    VCOMObject.OnKeyDown = Value
End Property

Property Get FlashWindowOnOpen() As Boolean
    FlashWindowOnOpen = VCOMObject.FlashWindowOnOpen
End Property

Property Let FlashWindowOnOpen(ByVal Value As Boolean)
    VCOMObject.FlashWindowOnOpen = Value
End Property

Property Get MoreInfoBackColor() As Long
    MoreInfoBackColor = VCOMObject.MoreInfoBackColor
End Property

Property Let MoreInfoBackColor(ByVal Value As Long)
    VCOMObject.MoreInfoBackColor = Value
End Property

Property Get AllowArrowKeys() As Boolean
    AllowArrowKeys = VCOMObject.AllowArrowKeys
End Property

Property Let AllowArrowKeys(ByVal Value As Boolean)
    VCOMObject.AllowArrowKeys = Value
End Property

Sub AddButton(ByVal DisplayName As String, ByVal ButtonAction As DefaultButtonActions)
    Call VCOMObject.AddButton(DisplayName, ButtonAction)
End Sub

Sub AddCustomButton(ByVal DisplayName As String, ByVal OnClickProcName As String)
    Call VCOMObject.AddCustomButton(DisplayName, OnClickProcName)
End Sub

Sub RemoveAllButtons()
    Call VCOMObject.RemoveAllButtons
End Sub

Function GetButtonIDByIndex(ByVal Index As Long) As Variant
    Call VCOMObject.AssignVar(GetButtonIDByIndex, VCOMObject.GetButtonIDByIndex(Index))
End Function

Function GetButtonCaptionByIndex(ByVal Index As Long) As Variant
    Call VCOMObject.AssignVar(GetButtonCaptionByIndex, VCOMObject.GetButtonCaptionByIndex(Index))
End Function

Property Get ButtonCount() As Long
    ButtonCount = VCOMObject.ButtonCount
End Property

Property Get Timeout() As Long
    Timeout = VCOMObject.Timeout
End Property

Property Let Timeout(ByVal Value As Long)
    VCOMObject.Timeout = Value
End Property

Property Get Timedout() As Boolean
    Timedout = VCOMObject.Timedout
End Property

Property Get CustomImageData() As String
    CustomImageData = VCOMObject.CustomImageData
End Property

Property Let CustomImageData(ByVal BitmapFileData As String)
    VCOMObject.CustomImageData = BitmapFileData
End Property

Sub LoadCustomImageFromFile(ByVal FilePath As String)
    Call VCOMObject.LoadCustomImageFromFile(FilePath)
End Sub

Function InternalCompressData(ByVal Data As String, ByVal CompressorModule As String) As String
    InternalCompressData = VCOMObject.InternalCompressData(Data, CompressorModule)
End Function

