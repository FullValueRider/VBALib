VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamArrayParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Provides the parsing of paramarrays for when we wish to interpret ParamArrays using one item or two item special cases"
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("VBALib")
'@ModuleDescription("Provides the parsing of paramarrays for when we wish to interpret ParamArrays using one item or two item special cases")

Private Const mcLocation        As String = "VBALib.ParamArrayParser"


Friend Function Parse(ByVal ipForwardedParamArray As Variant) As ParseResult
    
  Dim myresult As ParseResult
  Set myresult = ParseResult.Deb
  Set Parse = myresult
  
  Select Case Arrays.Count(ipForwardedParamArray)
    
    Case 0
    
      Exit Function
      
    Case 1
    
      Set myresult = Me.OneItem(ipForwardedParamArray(0))
      
      
    Case 2
      
      Set myresult = Me.TwoItems(ipForwardedParamArray(0), ipForwardedParamArray(1))
      
    
    Case Else
    
      Set myresult = Me.MultipleItems(ipForwardedParamArray)
      
 End Select
 
 Set Parse = myresult
 
End Function


' Easy one first
'@Description("Treats a ParamArray as just a ParamArray with no special cases")
Friend Function MultipleItems(ByVal ipForwardedParamArray As Variant) As ParseResult
Attribute MultipleItems.VB_Description = "Treats a ParamArray as just a ParamArray with no special cases"
  
    Dim myresult As ParseResult
    Set myresult = ParseResult.Deb
    Set MultipleItems = myresult
  
     With myresult
  
        .Code = IsItems
        .Items.AddValidatedIterable ipForwardedParamArray
      
    End With
  
End Function


'@Descscription("Interprets the single Item in the ParamArray.  Result is return in the Items property of ParseResult.)
Friend Function OneItem(ByVal ipItem As Variant) As ParseResult
        
  Dim myresult As ParseResult
  Set myresult = ParseResult.Deb
  Set OneItem = myresult
  
  Select Case True
          
      Case Types.Group.IsNotIterable(ipItem)
      
          myresult.Code = ParseResultCode.IsNonIterable
          myresult.Items.AddValidatedItem ipItem
      
          
      Case Arrays.IsMDArray(ipItem)
          ' We have no special iterable actions for arrays
          ' with more then two dimensions so add as single item
           
          myresult.Items.AddValidatedItem ipItem
          myresult.Code = ParseResultCode.IsMDArray
               
          
      'process a 2d array into a lyst of lysts
      Case Arrays.IsTableArray(ipItem)
              
          Set myresult.Items = Arrays.ToLystOfRankLysts(ipItem, RankIsRowFirstItemActionIsNoAction)
          myresult.Code = ParseResultCode.IsItems
          
          
      Case Types.Group.IsIterableItemsByEnum(ipItem)
      
          myresult.Items.AddValidatedIterable (ipItem)
          myresult.Code = ParseResultCode.IsItems
          
          
      Case Types.Group.IsIterableKeysByEnum(ipItem)
      
          myresult.Items.AddValidatedIterable (ipItem.Items)
          myresult.Keys.AddValidatedIterable (ipItem.Keys)
          myresult.Code = ParseResultCode.IsItemsAndKeys
      
          
      Case Else
      
          Guard ResultCode.UnexpectedItemInParamArray, True, mcLocation & "OneItem"
      
          
  End Select
        
End Function


Friend Function TwoItems(ByVal ipFirstItem As Variant, ByVal ipSecondItem As Variant) As ParseResult
    
    Guard ResultCode.InvalidType, Types.Group.IsAdmin(ipFirstItem), mcLocation & "TwoItemss", Array(TypeName(ipFirstItem)), "Admin types cannot be keys '{0}'"
    Dim myresult As ParseResult
    Set myresult = ParseResult.Deb
    Set TwoItems = myresult
      
    Select Case True
                
        Case Types.Group.IsNotIterable(ipFirstItem) 'Imlemented
    
            Set myresult = TwoItemsWhereFirstItemIsNotIterable(ipFirstItem, ipSecondItem)
            
            
        Case Arrays.IsMDArray(ipFirstItem)
        
            Guard ResultCode.InvalidIterable, True, mcLocation & "TwoItems"
            
            
        Case Arrays.IsTableArray(ipFirstItem)
            
            Set myresult = TwoItemsWhereFirstItemIsTableArray(ipFirstItem, ipSecondItem)
            
         
        Case Types.Group.IsIterableItemsByEnum(ipFirstItem) '
        
            Set myresult = TwoItemsWhereFirstItemIsIterableItemsByEnum(ipFirstItem, ipSecondItem)
            
            
        Case Types.Group.IsIterableKeysByEnum(ipFirstItem)
        
            Set myresult = TwoItemsWhereFirstItemIsIterableKeysByENum(ipFirstItem, ipSecondItem)
    
        
        Case Else
        
            Guard ResultCode.UnexpectedItemInParamArray, True, mcLocation & "TwoItems", TypeName(ipFirstItem)

        
    End Select
    
End Function


Friend Function TwoItemsWhereFirstItemIsNotIterable(ByVal ipFirstItem As Variant, ByVal ipSecondItem As Variant) As ParseResult
    
    Dim myresult As ParseResult
    Set myresult = ParseResult.Deb
    Set TwoItemsWhereFirstItemIsNotIterable = myresult
    
    Select Case True
        
        Case Types.Group.IsPrimitive(ipFirstItem), VBA.IsObject(ipFirstItem)
        
            myresult.Code = ParseResultCode.IsItemsAndKeys
            myresult.Items.AddValidatedItem ipFirstItem
            myresult.Items.AddValidatedItem ipSecondItem
        
        Case Else
        
            Guard ResultCode.UnexpectedItemInParamArray, True, mcLocation & "TwoItemsWHereFirstItemIsNotIterable", Array(TypeName(ipFirstItem))
        
    End Select
    
End Function


Friend Function TwoItemsWhereFirstItemIsTableArray(ByVal ipFirstItem As Variant, ByVal ipSecondItem As Variant) As ParseResult
    Select Case True
        
        
        
    End Select
End Function

Friend Function TwoItemsWhereFirstItemIsIterableItemsByEnum(ByVal ipFirstItem As Variant, ByVal ipSecondItem As Variant) As ParseResult
    Select Case True
        
        
        
    End Select
End Function

Friend Function TwoItemsWhereFirstItemIsIterableKeysByENum(ByVal ipFirstItem As Variant, ByVal ipSecondItem As Variant) As ParseResult
    Select Case True
        
        
        
    End Select
End Function
