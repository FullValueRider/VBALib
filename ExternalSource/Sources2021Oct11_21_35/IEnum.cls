VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
    
    Public Function Strings() As Variant
    End Function
    
    Public Function Members() As Variant
    End Function
    
    Public Function Actions() As Variant
    End Function
    
    Public Function Ord(ByVal ipEnum As GuardClause) As Long
    End Function
    
    Public Property Get AsEnum(ByVal ipEnum As GuardClause) As GuardClause
    End Property
    
    Public Function CanExec(ByVal ipEnum As GuardClause) As Boolean
    End Function
    
    Public Property Get ToString(ByVal ipEnum As GuardClause) As String
    End Property
        
    Public Property Get Exec(ByVal ipEnum As GuardClause, ByVal ipTableArray As Variant) As Object
    End Property
    
    Public Property Get IsMember(ByVal ipValue As Long) As Boolean
    End Property
    
    Public Property Get Name() As String
    End Property
    
    Public Property Get Count() As Long
    End Property
