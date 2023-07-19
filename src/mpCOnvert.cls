VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpCOnvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Mappers")
Option Explicit
Implements IMapper


Private Const MY_TYPENAME               As String = "mpConvert"


Private Type State
    Conversion                          As e_ConvertTo
End Type

Private s                               As State

Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property

'@Description("Converts the Items in the to the specifiied Type")
Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
Attribute IMapper_ExecMapper.VB_Description = "Converts the Items in the to the specifiied Type"
   
    On Error Resume Next
    Dim myResult As Variant
    Select Case s.Conversion
        
        Case e_ConvertTo.m_HexStr:          myResult = Hex$(ipHostItem)
        Case e_ConvertTo.m_OctStr:          myResult = Oct$(ipHostItem)
        Case e_ConvertTo.m_GroupName:       myResult = GroupInfo.Name(ipHostItem)
        Case e_ConvertTo.m_GroupId:         myResult = GroupInfo.Id(ipHostItem)
        Case e_ConvertTo.m_boolean:         myResult = VBA.CBool(ipHostItem)
        Case e_ConvertTo.m_Byte:            myResult = VBA.CByte(ipHostItem)
        Case e_ConvertTo.m_Currency:        myResult = VBA.CCur(ipHostItem)
        Case e_ConvertTo.m_Date::           myResult = VBA.CDate(ipHostItem)
        Case e_ConvertTo.m_Double:          myResult = VBA.CDbl(ipHostItem)
        Case e_ConvertTo.m_Decimal:         myResult = VBA.CDec(ipHostItem)
        Case e_ConvertTo.M_Integer:         myResult = VBA.CInt(ipHostItem)
        Case e_ConvertTo.m_Long:            myResult = VBA.CLng(ipHostItem)
        Case e_ConvertTo.m_LongLong:        myResult = VBA.CLngLng(ipHostItem)
        Case e_ConvertTo.m_LongPtr:         myResult = VBA.CLngPtr(ipHostItem)
        Case e_ConvertTo.m_Single:          myResult = VBA.CSng(ipHostItem)
        Case e_ConvertTo.m_string::         myResult = VBA.CStr(ipHostItem)
        
        Case Else
        
            Err.Raise _
                17 + vbObjectError, _
                "VBALib.mpConvertIMapper_ExecMapper", _
                "The requested conversion is not possible"
        
    End Select
    
    If Err.Number = 0 Then
        IMapper_ExecMapper = Array(myResult)
    Else
        IMapper_ExecMapper = Array(ipHostItem)
    End If
    
    On Error GoTo 0
    
End Function

'@DefaultMember
Public Function Deb(ByRef ipConversion As e_ConvertTo) As mpCOnvert
Attribute Deb.VB_UserMemId = 0
    With New mpCOnvert
       Set Deb = .ConstructInstance(ipConversion)
    End With
End Function

Friend Function ConstructInstance(ByRef ipConversion As e_ConvertTo) As mpCOnvert
    s.Conversion = ipConversion
    Set ConstructInstance = Me
End Function
