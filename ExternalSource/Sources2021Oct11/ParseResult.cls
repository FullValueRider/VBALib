VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParseResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Result from parsing the content using special casing when the paramarray is one or two items"
Option Explicit
'@Exposed
'@PredeclaredId

Public Enum ParseResultCode
    
	Failed = -1
	IsNonIterable
    IsMDArray
	IsItems
	IsItemsAndKeys

End Enum

Private Type Properties
	
	Code					As ParseResultCode
	NonIterableItem			As variant
	Items					As Lyst
	Keys					As Lyst
	
End Type


Private p			As Properties

Friend Function Deb _
( _
	Optional ipCode As ParseResultCode = failed, _
	Optional ipFailValue As Variant = Empty, _
	Optional ipItems As Lyst, _
	Optional ipKeys As Lyst _
) As ParseResult

	With New ParseResult
	
		Set Deb = .ReadytoUseInstance(ipCode, ipFailValue, ipItems, ipKeys)

	End With
	
End Function


Friend Function ReadytoUseInstance(ipCode As ParseResultCode, ipFailValue As Variant, ipItems As Lyst, ipKeys As Lyst) As ParseResult
	
	p.Code = ipCode
	
	If VBA.IsMissing(ipItems) Or (p.Items Is Nothing) Then
		
		Set p.Items = Lyst.Deb
		
	Else
	
		Set p.Items = ipItems
		
		
	End If
	
	If VBA.IsMissing(ipKeys) Or (p.Keys Is Nothing) Then
		
		Set p.Keys = Lyst.Deb
		
	Else
	
		Set p.Keys = ipItems
		
		
	End If
	
	
	If VBA.IsObject(ipFailValue) Then
		
		Set p.NonIterableItem = ipFailValue
		
		
	Else
		
		p.NonIterableItem = ipFailValue
		
		
	End If
	
	Set ReadytoUseInstance = Me
	
End Function

Public Property Get Code() As ParseResultCode
	Code = p.Code
End Property

Public Property Let Code(ByVal ipCode As ParseResultCode)
	p.Code = ipCode
End Property


Public Property Get NonIterableItem() As Variant
	
	If VBA.IsObject(p.NonIterableItem) Then
		
		Set NonIterableItem = p.NonIterableItem
		
		
	Else
		
		NonIterableItem = p.NonIterableItem
		
		
	End If
	
End Property

Public Property Set NonIterableItem(ByVal ipFailValue As Variant)
	
	If VBA.IsObject(ipFailValue) Then
		
		Set p.NonIterableItem = ipFailValue
		
		
	Else
		
		p.NonIterableItem = ipFailValue
		
		
	End If
	
End Property


Public Property Get Items() As Lyst
	Set Items = p.Items
End Property

Public Property Set Items(ByVal ipItems As Lyst)
	Set p.Items = ipItems
End Property


Public Property Get Keys() As Lyst
	Set Keys = p.Keys
End Property

Public Property Set Keys(ByVal ipKeys As Lyst)
	Set p.Keys = ipKeys
End Property