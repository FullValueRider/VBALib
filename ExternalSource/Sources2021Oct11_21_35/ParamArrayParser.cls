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
'@ModuleDescription("Provides the parsing of paramarrays for when we wish to interpret ParamArrays using one item or two item special cases")

Private Const mcLocation        As String = "VBALib.ParamArrayParser"


Friend Function Parse(ByVal ipForwardedParamArray As Variant) As ParseResult
    
  Dim myResult As ParseResult
  Set myResult = ParseResult.Deb
  Set Parse = myResult
  
  Select Case Arrays.Count(ipForwardedParamArray)
    
    Case 0
    
      Exit Function
      
    Case 1
    
      Set myResult = ParamArrayParser.OneItem(ipForwardedParamArray(0))
      
      
    Case 2
      
      Set myResult = ParamArrayParser.TwoItems(ipForwardedParamArray(0), ipForwardedParamArray(1))
      
    
    Case Else
    
      Set myResult = ParamArrayParser.MultipleItems(ipForwardedParamArray)
      
 End Select
 
 Set Parse = myResult
 
End Function


' Easy one first
'@Description("Treats a ParamArray as just a ParamArray with no special cases")
Friend Function MultipleItems(ByVal ipForwardedParamArray As Variant) As ParseResult
  
    Dim myResult As ParseResult
    Set myResult = ParseResult.Deb
    Set MultipleItems = myResult
  
     With myResult
  
        .Code = IsItems
        .Items.AddValidatedIterable ipForwardedParamArray
      
    End With
  
End Function


'@Descscription("Interprets the single Item in the ParamArray.  Result is return in the Items property of ParseResult.)
Friend Function OneItem(ByVal ipItem As Variant) As ParseResult
	
  Dim myResult As ParseResult
  Set myResult = ParseResult.Deb
  Set OneItem = myResult
  
  Select Case True
          
      Case Types.Group.IsNotIterable(ipItem)
      
          myResult.Code = ParseResultCode.IsNonIterable
          myResult.Items.AddValidatedItem ipItem
      
          
      Case Arrays.IsMDArray(ipItem)
          ' We have no special iterable actions for arrays 
          ' with more then two dimensions so add as single item
           
          myResult.Items.AddValidatedItem ipItem
          myResult.Code = ParseResultCode.IsMDArray
               
          
      'process a 2d array into a lyst of lysts        
      Case Arrays.IsTableArray(ipItem)
              
          Set myResult.Items = Arrays.ToLystOfRankLysts(ipItem, RankIsRowFirstItemActionIsNoAction)
          myResult.Code = ParseResultCode.IsItems
          
          
      Case Types.Group.IsIterableItemsByEnum(ipItem)
      
          myResult.Items.AddValidatedIterable(ipItem)
          myResult.Code = ParseResultCode.IsItems
          
          
      Case Types.Group.IsIterableKeysByEnum(ipItem)
      
          myResult.Items.AddValidatedIterable(ipItem.Items)
          myResult.Keys.AddValidatedIterable(ipItem.Keys)
          myResult.Code = ParseResultCode.IsItemsAndKeys
      
          
      Case Else
      
          Guard GuardClause.UnexpectedItemInParamArray, True, mcLocation & "OneItem"
      
          
  End Select
        
End Function


Friend Function TwoItems(ByVal ipFirstItem As Variant, ByVal ipSecondItem As Variant) As ParseResult
    
    Guard GuardClause.InvalidType, Types.Group.IsAdmin(ipFirstItem), mcLocation & "TwoItemss", Array(TypeName(ipFirstItem)), "Admin types cannot be keys '{0}'"
    Dim myResult As ParseResult
    Set myResult = ParseResult.Deb
    Set TwoItems = myResult
      
    Select Case True
                
        Case Types.Group.IsNotIterable(ipFirstItem) 'Imlemented
    
            Set myResult = TwoItemsWhereFirstItemIsNotIterable(ipFirstItem, ipSecondItem)
            
            
        Case Arrays.IsMDArray(ipFirstItem)
        
            Guard GuardClause.InvalidIterable, True, mcLocation & "TwoItems"
            
            
        Case Arrays.IsTableArray(ipFirstItem)
            
            Set myResult = TwoItemsWhereFirstItemIsTableArray(ipFirstItem, ipSecondItem)
            
         
        Case Types.Group.IsIterableItemsByEnum(ipFirstItem) '
        
            Set myResult = TwoItemsWhereFirstItemIsIterableItemsByEnum(ipFirstItem, ipSecondItem)
            
            
        Case Types.Group.IsIterableKeysByEnum(ipFirstItem)
        
            Set myResult = TwoItemsWhereFirstItemIsIterableKeysByENum(ipFirstItem, ipSecondItem)
    
        
        Case Else
        
            Guard GuardClause.UnexpectedItemInParamArray, True, mcLocation & "TwoItems", TypeName(ipFirstItem)

        
    End Select
    
End Function


Friend Function TwoItemsWhereFirstItemIsNotIterable(ByVal ipFirstItem As Variant, ByVal ipSecondItem As Variant) As ParseResult
    
    Dim myResult As ParseResult
    Set myResult = ParseResult.Deb
    Set TwoItemsWhereFirstItemIsNotIterable = myResult
    
    Select Case True
        
        Case Types.Group.IsPrimitive(ipFirstItem), VBA.IsObject(ipFirstItem)
        
            myResult.Code = ParseResultCode.IsItemsAndKeys
            myResult.Items.AddValidatedItem ipFirstItem
            myResult.Items.AddValidatedItem ipSecondItem
        
        Case Else
        
            Guard GuardClause.UnexpectedItemInParamArray, True, mcLocation & "TwoItemsWHereFirstItemIsNotIterable", Array(TypeName(ipFirstItem))
        
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