Public Enum Action
	
	Equal
	NotEqual
	LessThan
	LessThanOrEqual
	NotMoreThan
	MoreThan
	MoreThanOrEqual
	NotLessThan
	
End Enum

Public Function Compare(ByVal ipTest As Action, ByVal ipVar1 As Variant, ByVal ipvar2 As Variant) As Boolean
	
	Guard GuardClause.NotSameType, TypeName(ipVar1) <> TypeName(ipvar2), "VBALib.Comparer.Compare", Array(TypeName(ipVar1), TypeName(ipvar2))
	
	Dim myResult As Boolean
	
	Select Case ipTest
		
		Case Action.Equal
			
			myResult = ipVar1 = ipvar2
		
			
		Case Action.NotEqual
		
			myResult = ipVar1 <> ipvar2
			
			
		Case Action.LessThan
		
			myResult = ipVar1 < ipvar2
			
			
		Case LessThanOrEqual, NotMoreThan
		
			myResult = ipVar1 <= ipvar2
		
			
		Case MoreThan
		
			myResult = ipVar1 > ipvar2
			
			
		Case MoreThanOrEqual, NotLessThan
		
			myResult = ipVar1 >= ipvar2
		
		
	End Select
	
	Compare = myResult
	
End Function

	

