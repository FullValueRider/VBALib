Attribute VB_Name = "Guarding"
    '@Folder("Guarding")
    Option Explicit
    
'    Public Const UnexpectedResult                   As Boolean = True
'    Public Const NoGuardTextParams                  As Variant = Empty
    
    Public Sub Guard _
    ( _
        ByVal ipGuardId As ResultCode, _
        ByVal ipThrow As Boolean, _
        ByVal ipLocation As String, _
        Optional ByVal ipArgs As Variant = Empty, _
        Optional ByVal ipAltMessage As String = vbNullString _
    )
            
        If ipThrow Then
            
            Dim myargs As Variant
            If Arrays.HasItems(ipArgs) Then
                
                myargs = ipArgs
                
            Else
                
                myargs = Array(ipArgs)
                
            End If
            
            Dim myMessage As String
            If VBA.Len(ipAltMessage) = 0 Then
            
                myMessage = Enums.GuardClauses.ToString(ipGuardId)
                
            Else
                
                myMessage = ipAltMessage
                
            End If
            
            If Arrays.HasItems(ipArgs) Then
                
                myMessage = Fmt.TxtArr(myMessage, myargs)
                
            End If

            VBA.Information.Err.Raise ipGuardId, ipLocation, myMessage
                        
        End If
        
    End Sub


