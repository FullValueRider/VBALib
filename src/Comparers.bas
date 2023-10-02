Attribute VB_Name = "Comparers"
'@Folder("Helpers")
Option Explicit
' **Length vs content
' Comparison of strings in VBA shows that the default behaviour
' Is based on content then length

' To avoid discombobulation the Comparers below should follow the same rule.

' We also need to be aware that in some cases comparing against types
' classes as admin is a legitimate comparison, e.g. Nothing is an acceptable
' comparison for  any object.

' There are also two additional considerations
' Do comparisons need to be type specific within a type group
' i.e. Integer2 is not the same as long 2
' for container classes does the comparison need to respect the order of items
' i.e. [2,3] is not the same as [3,2].

' Finally, when using fmt.Text to obtain string representations of objects for comparision purposed
' we need to make sure that markup is not used

' A word about admin types
' empty coerces to 0 for numbers and vb

'***ToDO Updat Fmt text to have a version that includes types.  This will simplify the continer comparers.

' We also need to be aware
Public Function EQ(ByRef ipLHS As Variant, ByRef ipRhs As Variant, Optional ByRef ipTypes As Boolean = False, Optional ByRef ipOrder As Boolean = True, Optional ByRef ipMismatchIsFalse As Boolean = True) As Boolean
    
    ' If type comparison is required the test can exit early if the Types Do not match
    If ipTypes Then
        If VBA.TypeName(ipLHS) <> VBA.TypeName(ipRhs) Then
            EQ = MismatchOrderer(ipLHS) = MismatchOrderer(ipRhs)
            Exit Function
        End If
    End If
    
    ' A value cannot be equal to nothing
    If GroupInfo.IsAdmin(ipLHS) Xor GroupInfo.IsAdmin(ipRhs) Then
        EQ = False
        
    ' all admin values are considered equal
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        EQ = VBA.TypeName(ipLHS) = VBA.TypeName(ipRhs)
        
    ElseIf GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRhs) Then
        EQ = ipLHS = ipRhs
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRhs) Then
            EQ = ipLHS = ipRhs
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRhs) Then
        EQ = ipLHS = ipRhs
        
    ElseIf GroupInfo.IsContainer(ipLHS) Or GroupInfo.IsContainer(ipRhs) Then
        EQ = ContainersEQ(ipLHS, ipRhs, ipTypes, ipOrder, ipMismatchIsFalse)
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRhs) Then
        EQ = Fmt.NoMarkup.Text("{0}", ipLHS) = Fmt.NoMarkup.Text("{0}", ipRhs)
       
    ' User is trying to compare two different types
    ' the Comparers do not support type cooercion
    ' its good for the users soul.
    Else
        If ipMismatchIsFalse Then
            EQ = False 'MismatchOrderer(ipLHS) = MismatchOrderer(ipRHS)
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If
    
End Function
    
Private Function ContainersEQ(ByRef ipLHS As Variant, ByRef ipRhs As Variant, ByRef ipTypes As Boolean, ByRef ipOrder As Boolean, ByRef ipMismatchIsFalse As Boolean) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRhs)

    ' VBA returns true for 'Nothing is Nothing'
    If IsNothing(ipLHS) And IsNothing(ipRhs) Then
        ContainersEQ = False
        
    ' If one object is nothing they cannot be equal
    ElseIf IsNothing(ipLHS) Or IsNothing(ipRhs) Then
        ContainersEQ = False
        
        
    ElseIf GroupInfo.IsItemByKey(ipLHS) And GroupInfo.IsItemByKey(ipRhs) Then
    
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
        And (GroupInfo.IsList(ipRhs) Or GroupInfo.IsArray(ipRhs)) _
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
    
        If ipMismatchIsFalse Then
            ContainersEQ = False
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If
    
End Function

Public Function NEQ(ByRef ipLHS As Variant, ByRef ipRhs As Variant, Optional ByRef ipTypes As Boolean = False, Optional ByRef ipOrder As Boolean = True, Optional ByRef ipMismatchIsFalse As Boolean = True) As Boolean
    NEQ = Not EQ(ipLHS, ipRhs, ipTypes, ipOrder, ipMismatchIsFalse)
End Function


Public Function MT(ByRef ipLHS As Variant, ByRef ipRhs As Variant, Optional ByRef ipTypes As Boolean = False, Optional ByRef ipOrder As Boolean = True, Optional ByRef ipMismatchIsFalse As Boolean = True) As Boolean
    
    If ipTypes Then
        If VBA.TypeName(ipLHS) <> VBA.TypeName(ipRhs) Then
            MT = MismatchOrderer(ipLHS) > MismatchOrderer(ipRhs)
            Exit Function
        End If
    End If

    If GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        MT = False
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsNotAdmin(ipRhs) Then
        MT = False
    ElseIf GroupInfo.IsNotAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        MT = True
        
    ElseIf GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRhs) Then
        MT = False
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRhs) Then
            MT = ipLHS > ipRhs
        
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRhs) Then
        MT = ipLHS > ipRhs
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRhs) Then
        MT = Fmt.NoMarkup.Text("{0}", ipLHS) > Fmt.NoMarkup.Text("{0}", ipRhs)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRhs) Then
        MT = ContainersMT(ipLHS, ipRhs, ipTypes, ipOrder, ipMismatchIsFalse)
        
    Else
    
        If ipMismatchIsFalse Then
            MT = False 'MismatchOrderer(ipLHS) > MismatchOrderer(ipRHS)
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If
    
End Function
    
Private Function ContainersMT(ByRef ipLHS As Variant, ByRef ipRhs As Variant, ByRef ipTypes As Boolean, ByRef ipOrder As Boolean, ByRef ipMismatchIsFalse As Boolean) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRhs)
        
    If GroupInfo.IsItemByKey(ipLHS) And GroupInfo.IsItemByKey(ipRhs) Then
    
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
        And (GroupInfo.IsList(ipRhs) Or GroupInfo.IsArray(ipRhs)) _
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
    
        If ipMismatchIsFalse Then
            ContainersMT = False
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If
    
End Function

'Public Function MTEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
'    MTEQ = Not LT(ipLHS, ipRHS)
'End Function

Public Function MTEQ(ByRef ipLHS As Variant, ByRef ipRhs As Variant, Optional ByRef ipTypes As Boolean = False, Optional ByRef ipOrder As Boolean = True, Optional ByRef ipMismatchIsFalse As Boolean = True) As Boolean

    If ipTypes Then
        If VBA.TypeName(ipLHS) <> VBA.TypeName(ipRhs) Then
            MTEQ = MismatchOrderer(ipLHS) >= MismatchOrderer(ipRhs)
            Exit Function
        End If
    End If


    If GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        MTEQ = True
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsNotAdmin(ipRhs) Then
        MTEQ = False
    ElseIf GroupInfo.IsNotAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        MTEQ = True

    ElseIf GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRhs) Then
        MTEQ = ipLHS = ipRhs

    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRhs) Then
        MTEQ = ipLHS >= ipRhs

    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRhs) Then
        MTEQ = ipLHS >= ipRhs

    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRhs) Then
        MTEQ = Fmt.NoMarkup.Text("{0}", ipLHS) >= Fmt.NoMarkup.Text("{0}", ipRhs)

    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRhs) Then
        MTEQ = ContainersMTEQ(ipLHS, ipRhs, ipTypes, ipOrder, ipMismatchIsFalse)

    Else
            
        If ipMismatchIsFalse Then
            MTEQ = False 'MismatchOrderer(ipLHS) >= MismatchOrderer(ipRHS)
        Else
            TypeMismatch ipLHS, ipRhs
        End If

    End If

End Function

Private Function ContainersMTEQ(ByRef ipLHS As Variant, ByRef ipRhs As Variant, ByRef ipTypes As Boolean, ByRef ipOrder As Boolean, ByRef ipMismatchIsFalse As Boolean) As Boolean

    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRhs)

    If GroupInfo.IsItemByKey(ipLHS) And GroupInfo.IsItemByKey(ipRhs) Then

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
        And (GroupInfo.IsList(ipRhs) Or GroupInfo.IsArray(ipRhs)) _
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
    
        If ipMismatchIsFalse Then
            ContainersMTEQ = False
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If

End Function

Public Function LT(ByRef ipLHS As Variant, ByRef ipRhs As Variant, Optional ByRef ipTypes As Boolean = False, Optional ByRef ipOrder As Boolean = True, Optional ByRef ipMismatchIsFalse As Boolean = True) As Boolean
    
    If ipTypes Then
        If VBA.TypeName(ipLHS) <> VBA.TypeName(ipRhs) Then
            LT = MismatchOrderer(ipLHS) < MismatchOrderer(ipRhs)
            Exit Function
        End If
    End If

    
    If GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        LT = False
        
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsNotAdmin(ipRhs) Then
        LT = True
        
    ElseIf GroupInfo.IsNotAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        LT = False

    ElseIf GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRhs) Then
        LT = False
        
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRhs) Then
        LT = ipLHS < ipRhs

    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRhs) Then
        LT = ipLHS < ipRhs
    
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        LT = False
   
    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRhs) Then
        LT = Fmt.NoMarkup.Text("{0}", ipLHS) < Fmt.NoMarkup.Text("{0}", ipRhs)
       
    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRhs) Then
        LT = ContainersLT(ipLHS, ipRhs, ipTypes, ipOrder, ipMismatchIsFalse)
        
    Else
    
        If ipMismatchIsFalse Then
            LT = False  'MismatchOrderer(ipLHS) < MismatchOrderer(ipRHS)
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If
    
End Function
    
Private Function ContainersLT(ByRef ipLHS As Variant, ByRef ipRhs As Variant, ByRef ipTypes As Boolean, ByRef ipOrder As Boolean, ByRef ipMismatchIsFalse As Boolean) As Boolean
    
    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRhs)
        
    If GroupInfo.IsItemByKey(ipLHS) And GroupInfo.IsItemByKey(ipRhs) Then
    
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
        And (GroupInfo.IsList(ipRhs) Or GroupInfo.IsArray(ipRhs)) _
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
    
        If ipMismatchIsFalse Then
            ContainersLT = False
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If
    
End Function

'Public Function LTEQ(ByRef ipLHS As Variant, ByRef ipRHS As Variant) As Boolean
'    LTEQ = Not MT(ipLHS, ipRHS)
'End Function

Public Function LTEQ(ByRef ipLHS As Variant, ByRef ipRhs As Variant, Optional ByRef ipTypes As Boolean = False, Optional ByRef ipOrder As Boolean = True, Optional ByRef ipMismatchIsFalse As Boolean = True) As Boolean

    If ipTypes Then
        If VBA.TypeName(ipLHS) <> VBA.TypeName(ipRhs) Then
            LTEQ = MismatchOrderer(ipLHS) <= MismatchOrderer(ipRhs)
            Exit Function
        End If
    End If

    
    If GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        LTEQ = True
    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsNotAdmin(ipRhs) Then
        LTEQ = True
    ElseIf GroupInfo.IsNotAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        LTEQ = False

    ElseIf GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRhs) Then
        LTEQ = ipLHS = ipRhs

    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRhs) Then
        LTEQ = ipLHS <= ipRhs

    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRhs) Then
        LTEQ = ipLHS <= ipRhs

    ElseIf GroupInfo.IsAdmin(ipLHS) And GroupInfo.IsAdmin(ipRhs) Then
        LTEQ = Fmt.NoMarkup.Text("{0}", ipLHS) = Fmt.NoMarkup.Text("{0}", ipRhs)

    ElseIf GroupInfo.IsItemObject(ipLHS) And GroupInfo.IsItemObject(ipRhs) Then
        LTEQ = Fmt.NoMarkup.Text("{0}", ipLHS) <= Fmt.NoMarkup.Text("{0}", ipRhs)

    ElseIf GroupInfo.IsContainer(ipLHS) And GroupInfo.IsContainer(ipRhs) Then
        LTEQ = ContainersLTEQ(ipLHS, ipRhs, ipTypes, ipOrder, ipMismatchIsFalse)

    Else
    
        If ipMismatchIsFalse Then
            LTEQ = False '  MismatchOrderer(ipLHS) <= MismatchOrderer(ipRHS)
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If

End Function

Private Function ContainersLTEQ(ByRef ipLHS As Variant, ByRef ipRhs As Variant, ByRef ipTypes As Boolean, ByRef ipOrder As Boolean, ByRef ipMismatchIsFalse As Boolean) As Boolean

    Dim myLItems As IterItems: Set myLItems = IterItems(ipLHS)
    Dim myRItems As IterItems: Set myRItems = IterItems(ipRhs)

    If GroupInfo.IsItemByKey(ipLHS) And GroupInfo.IsItemByKey(ipRhs) Then

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
        And (GroupInfo.IsList(ipRhs) Or GroupInfo.IsArray(ipRhs)) _
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
    
        If ipMismatchIsFalse Then
            ContainersLTEQ = False
        Else
            TypeMismatch ipLHS, ipRhs
        End If
        
    End If

End Function

Private Sub TypeMismatch(ByRef ipLHS As Variant, ByRef ipRhs As Variant)

    Err.Raise 17 + vbObjectError, _
        "VBALib.Comparer", _
        Fmt.Text("ipLHS was {0}:{1}, ipRHS was {2}:{3).", VBA.TypeName(ipLHS), ipLHS, VBA.TypeName(ipRhs), ipRhs)
        
End Sub


' Comparer needs a method to resolve problems associated with returning false
' when types do not match e.g. both LT and MT return false.
' this means that code elsewhere can never resolve
' so we need to set a prority for different types so that MT/LT can return true or false
' To do this we assign a sinle numerical value to a TypeGroup
' so that Lt or MT can return a true or false.
' The assignments are
' Admin = 0
' Boolean = 1
' Number = 2
' String = 3
' ItemObject = 4
' Array = 5
' Container = 6

    
            
Private Function MismatchOrderer(ByRef ipItem As Variant) As Long
    Select Case True
        Case GroupInfo.IsAdmin(ipItem)
            MismatchOrderer = 0
        Case GroupInfo.IsBoolean(ipItem)
            MismatchOrderer = 1
        Case GroupInfo.IsNumber(ipItem)
            MismatchOrderer = 2
        Case GroupInfo.IsString(ipItem)
            MismatchOrderer = 3
        Case GroupInfo.IsItemObject(ipItem)
            MismatchOrderer = 4
        Case GroupInfo.IsArray(ipItem)
            MismatchOrderer = 5
        Case GroupInfo.IsContainer(ipItem)
            MismatchOrderer = 6
    End Select
End Function

