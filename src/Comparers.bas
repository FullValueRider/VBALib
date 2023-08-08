Attribute VB_Name = "Comparers"
'@Folder("Helpers")
Option Explicit

' Comparisons of ipLHS and ipRHS are based on the following priorities
' 1. ipLHS and ipRHS must be in the same group
' 2. Containers objects must have the same number of items to proceed to a comparison of content
' 3. Dictionaries are compared in order of Key/Value pairs.
' 4. Lists and Arrays are compared based on order of items
' 5. Admin types are compared based on their string representation
' 6. Strings must be the same length to proceed to a comparison of content
' 7. Comparison of string content is according to the current Option Compare setting
' 7. Boolean and Admin items can only ever be EQ or NEQ

Public Function EQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    If GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRHS) Then
        EQ = ipLHS = ipRHS
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRHS) Then
        If VBA.Len(ipLHS) <> VBA.Len(ipRHS) Then
            EQ = False
        Else
            EQ = ipLHS = ipRHS
        End If
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRHS) Then
        EQ = ipLHS = ipRHS
    
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRHS) Then
        EQ = Fmt.NoMarkup.Text("{0}", ipLHS) = Fmt.NoMarkup.Text("{0}", ipRHS)
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRHS) Then
        EQ = Fmt.NoMarkup.Text("{0}", ipLHS) = Fmt.NoMarkup.Text("{0}", ipRHS)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRHS) Then
        EQ = ContainersEQ(ipLHS, ipRHS)
    
    Else
        EQ = False
        
    End If
    
End Function
    
Private Function ContainersEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRHS)

        
    If GroupInfo.IsDictionary(ipLHS) And GroupInfo.IsDictionary(ipRHS) Then
    
       If myLItems.Size <> myRItems.Size Then
            ContainersEQ = False
            Exit Function
        End If

        Do

            If NEQ(myLItems.CurKey(0), myRItems.CurKey(0)) Then
                ContainersEQ = False
                Exit Function
            End If
            
            If NEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersEQ = False
                Exit Function
            End If
                
        Loop While myLItems.MoveNext And myRItems.MoveNext
       
        ContainersEQ = True
            
    ElseIf _
        (GroupInfo.IsList(ipLHS) Or GroupInfo.IsArray(ipLHS)) _
        And (GroupInfo.IsList(ipRHS) Or GroupInfo.IsArray(ipRHS)) _
    Then

        If myLItems.Size <> myRItems.Size Then
            ContainersEQ = False
            Exit Function
        End If
    
        Do
            If NEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersEQ = False
                Exit Function
            End If
        Loop While myLItems.MoveNext And myRItems.MoveNext
        
        ContainersEQ = True
        
    Else
    
        ContainersEQ = False
        
    End If
    
End Function

Public Function NEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    If GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRHS) Then
        NEQ = ipLHS <> ipRHS
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRHS) Then
        If VBA.Len(ipLHS) <> VBA.Len(ipRHS) Then
            NEQ = True
        Else
            NEQ = ipLHS <> ipRHS
        End If
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRHS) Then
        NEQ = ipLHS <> ipRHS
    
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRHS) Then
        NEQ = Fmt.NoMarkup.Text("{0}", ipLHS) <> Fmt.NoMarkup.Text("{0}", ipRHS)
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRHS) Then
        NEQ = Fmt.NoMarkup.Text("{0}", ipLHS) <> Fmt.NoMarkup.Text("{0}", ipRHS)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRHS) Then
        NEQ = ContainersNEQ(ipLHS, ipRHS)
        
    Else
    
        NEQ = True
        
    End If
    
End Function
    
Private Function ContainersNEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRHS)

    If myLItems.Size <> myRItems.Size Then
        ContainersNEQ = True
        Exit Function
    End If
        
    If GroupInfo.IsDictionary(ipLHS) And GroupInfo.IsDictionary(ipRHS) Then

        If myLItems.Size <> myRItems.Size Then
            ContainersNEQ = True
            Exit Function
        End If

        Do

            If NEQ(myLItems.CurKey(0), myRItems.CurKey(0)) Then
                ContainersNEQ = True
                Exit Function
            End If
            
            If NEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersNEQ = True
                Exit Function
            End If
                
        Loop While myLItems.MoveNext And myRItems.MoveNext
       
        ContainersNEQ = False
            
    ElseIf _
        (GroupInfo.IsList(ipLHS) Or GroupInfo.IsArray(ipLHS)) _
        And (GroupInfo.IsList(ipRHS) Or GroupInfo.IsArray(ipRHS)) _
    Then
    
        If myLItems.Size <> myRItems.Size Then
            ContainersNEQ = True
            Exit Function
        End If

        Do
            If NEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersNEQ = True
                Exit Function
            End If
        Loop While myLItems.MoveNext And myRItems.MoveNext
        
        ContainersNEQ = False
        
    Else
    
        ContainersNEQ = True
        
    End If
    
End Function

Public Function MT(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    If GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRHS) Then
        MT = False
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRHS) Then
        If VBA.Len(ipLHS) > VBA.Len(ipRHS) Then
            MT = True
        ElseIf VBA.Len(ipLHS) < VBA.Len(ipRHS) Then
            MT = False
        Else
            MT = ipLHS > ipRHS
        End If
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRHS) Then
        MT = ipLHS > ipRHS
    
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRHS) Then
        MT = False
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRHS) Then
        MT = Fmt.NoMarkup.Text("{0}", ipLHS) > Fmt.NoMarkup.Text("{0}", ipRHS)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRHS) Then
        MT = ContainersMT(ipLHS, ipRHS)
        
    Else
        MT = False
        
    End If
    
End Function
    
Private Function ContainersMT(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRHS)
        
    If GroupInfo.IsDictionary(ipLHS) And GroupInfo.IsDictionary(ipRHS) Then
    
        If myLItems.Size > myRItems.Size Then
            ContainersMT = True
            Exit Function
        ElseIf myLItems.Size < myRItems.Size Then
            ContainersMT = False
            Exit Function
        End If

        Do

            If MT(myLItems.CurKey(0), myRItems.CurKey(0)) Then
                ContainersMT = True
                Exit Function
            End If
            
            If MT(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersMT = True
                Exit Function
            End If
                
        Loop While myLItems.MoveNext And myRItems.MoveNext
       
        ContainersMT = False
            
    ElseIf _
        (GroupInfo.IsList(ipLHS) Or GroupInfo.IsArray(ipLHS)) _
        And (GroupInfo.IsList(ipRHS) Or GroupInfo.IsArray(ipRHS)) _
    Then
    
        If myLItems.Size > myRItems.Size Then
            ContainersMT = True
            Exit Function
        ElseIf myLItems.Size < myRItems.Size Then
            ContainersMT = False
            Exit Function
        End If
    
        Do
            If MT(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersMT = True
                Exit Function
            End If
        Loop While myLItems.MoveNext And myRItems.MoveNext
        
        ContainersMT = False
        
    Else
    
        ContainersMT = False
        
    End If
    
End Function

Public Function MTEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    If GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRHS) Then
        MTEQ = ipLHS = ipRHS
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRHS) Then
        If VBA.Len(ipLHS) > VBA.Len(ipRHS) Then
            MTEQ = True
        ElseIf VBA.Len(ipLHS) < VBA.Len(ipRHS) Then
            MTEQ = False
        Else
            MTEQ = ipLHS >= ipRHS
        End If
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRHS) Then
        MTEQ = ipLHS >= ipRHS
    
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRHS) Then
        MTEQ = Fmt.NoMarkup.Text("{0}", ipLHS) = Fmt.NoMarkup.Text("{0}", ipRHS)
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRHS) Then
        MTEQ = Fmt.NoMarkup.Text("{0}", ipLHS) >= Fmt.NoMarkup.Text("{0}", ipRHS)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRHS) Then
        MTEQ = ContainersMTEQ(ipLHS, ipRHS)
        
    Else
        MTEQ = False
        
    End If
    
End Function
    
Private Function ContainersMTEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRHS)

    If GroupInfo.IsDictionary(ipLHS) And GroupInfo.IsDictionary(ipRHS) Then
    
        If myLItems.Size < myRItems.Size Then
            ContainersMTEQ = False
            Exit Function
        ElseIf myLItems.Size > myRItems.Size Then
            ContainersMTEQ = True
        End If
    
        Do

            If MTEQ(myLItems.CurKey(0), myRItems.CurKey(0)) Then
                ContainersMTEQ = True
                Exit Function
            End If
            
            If MTEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersMTEQ = True
                Exit Function
            End If
                
        Loop While myLItems.MoveNext And myRItems.MoveNext
       
        ContainersMTEQ = False
            
    ElseIf _
        (GroupInfo.IsList(ipLHS) Or GroupInfo.IsArray(ipLHS)) _
        And (GroupInfo.IsList(ipRHS) Or GroupInfo.IsArray(ipRHS)) _
    Then
    
        If myLItems.Size < myRItems.Size Then
            ContainersMTEQ = False
            Exit Function
        ElseIf myLItems.Size > myRItems.Size Then
            ContainersMTEQ = True
        End If

        Do
            If MTEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersMTEQ = True
                Exit Function
            End If
        Loop While myLItems.MoveNext And myRItems.MoveNext
        
        ContainersMTEQ = False
        
    Else
    
        ContainersMTEQ = False
        
    End If
    
End Function

Public Function LT(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    If GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRHS) Then
        LT = False
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRHS) Then
        If VBA.Len(ipLHS) > VBA.Len(ipRHS) Then
            LT = False
        ElseIf VBA.Len(ipLHS) < VBA.Len(ipRHS) Then
            LT = True
        Else
            LT = ipLHS < ipRHS
        End If
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRHS) Then
        LT = ipLHS < ipRHS
    
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRHS) Then
        LT = False
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRHS) Then
        LT = Fmt.NoMarkup.Text("{0}", ipLHS) < Fmt.NoMarkup.Text("{0}", ipRHS)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRHS) Then
        LT = ContainersLT(ipLHS, ipRHS)
        
    Else
    
        LT = False
        
    End If
    
End Function
    
Private Function ContainersLT(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRHS)
        
    If GroupInfo.IsDictionary(ipLHS) And GroupInfo.IsDictionary(ipRHS) Then
    
        If myLItems.Size < myRItems.Size Then
            ContainersLT = True
            Exit Function
        ElseIf myLItems.Size > myRItems.Size Then
            ContainersLT = False
            Exit Function
        End If
    
        Do

            If LT(myLItems.CurKey(0), myRItems.CurKey(0)) Then
                ContainersLT = True
                Exit Function
            End If
            
            If LT(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersLT = True
                Exit Function
            End If
                
        Loop While myLItems.MoveNext And myRItems.MoveNext
       
        ContainersLT = False
            
    ElseIf _
        (GroupInfo.IsList(ipLHS) Or GroupInfo.IsArray(ipLHS)) _
        And (GroupInfo.IsList(ipRHS) Or GroupInfo.IsArray(ipRHS)) _
    Then
    
        If myLItems.Size < myRItems.Size Then
        ContainersLT = True
            Exit Function
        ElseIf myLItems.Size > myRItems.Size Then
            ContainersLT = False
            Exit Function
        End If
    
        Do
            If LT(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersLT = True
                Exit Function
            End If
        Loop While myLItems.MoveNext And myRItems.MoveNext
        
        ContainersLT = False
        
    Else
    
        ContainersLT = False
        
    End If
    
End Function

Public Function LTEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    If GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRHS) Then
        LTEQ = ipLHS = ipRHS
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRHS) Then
        If VBA.Len(ipLHS) > VBA.Len(ipRHS) Then
            LTEQ = False
        ElseIf VBA.Len(ipLHS) < VBA.Len(ipRHS) Then
            LTEQ = True
        Else
            LTEQ = ipLHS <= ipRHS
        End If
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRHS) Then
        LTEQ = ipLHS <= ipRHS
    
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRHS) Then
        LTEQ = Fmt.NoMarkup.Text("{0}", ipLHS) = Fmt.NoMarkup.Text("{0}", ipRHS)
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRHS) Then
        LTEQ = Fmt.NoMarkup.Text("{0}", ipLHS) <= Fmt.NoMarkup.Text("{0}", ipRHS)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRHS) Then
        LTEQ = ContainersLTEQ(ipLHS, ipRHS)
        
    Else
        LTEQ = False
        
    End If
    
End Function
    
Private Function ContainersLTEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRHS)

    If GroupInfo.IsDictionary(ipLHS) And GroupInfo.IsDictionary(ipRHS) Then
    
        If myLItems.Size < myRItems.Size Then
            ContainersLTEQ = True
            Exit Function
        ElseIf myLItems.Size > myRItems.Size Then
            ContainersLTEQ = False
            Exit Function
        End If

        Do

            If LTEQ(myLItems.CurKey(0), myRItems.CurKey(0)) Then
                ContainersLTEQ = True
                Exit Function
            End If
            
            If LTEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersLTEQ = True
                Exit Function
            End If
                
        Loop While myLItems.MoveNext And myRItems.MoveNext
       
        ContainersLTEQ = False
            
    ElseIf _
        (GroupInfo.IsList(ipLHS) Or GroupInfo.IsArray(ipLHS)) _
        And (GroupInfo.IsList(ipRHS) Or GroupInfo.IsArray(ipRHS)) _
    Then

        If myLItems.Size < myRItems.Size Then
            ContainersLTEQ = True
            Exit Function
        ElseIf myLItems.Size > myRItems.Size Then
            ContainersLTEQ = False
            Exit Function
        End If
        

        Do
            If LTEQ(myLItems.CurItem(0), myRItems.CurItem(0)) Then
                ContainersLTEQ = True
                Exit Function
            End If
        Loop While myLItems.MoveNext And myRItems.MoveNext
        
        ContainersLTEQ = False
        
    Else
    
        ContainersLTEQ = False
        
    End If
    
End Function
