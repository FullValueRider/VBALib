Attribute VB_Name = "Enums"
Attribute VB_Description = "A location for enums used by multiple classes"
'@Folder("Constants")
Option Explicit

'@ModuleDescription("A location for enums used by multiple classes")

Public Enum e_ConvertTo
    m_First = 0
    m_HexStr = m_First
    m_OctStr
    m_GroupName
    m_GroupId
    m_LongPtr
    m_boolean
    m_Byte
    m_Currency
    m_Date
    m_Double
    m_Decimal
    M_Integer
    m_Long
    m_LongLong
    m_Single
    m_string
    m_Last = m_string
End Enum


Public Enum e_MirrorType
    m_First = 0
    m_ByAllValues = m_First
    m_ByFirstValue
    m_Last = m_ByFirstValue
End Enum


Public Enum e_SetoF
    m_First = 0
    m_Common = m_First
    m_HostOnly
    m_ParamOnly
    m_NotCommon
    m_Unique
    m_Last = m_NotCommon
End Enum


Public Enum e_Get
    m_First = 0
    m_Unique = m_First
    m_All
    m_Last = m_All
End Enum


