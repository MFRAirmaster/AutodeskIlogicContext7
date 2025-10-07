' Title: Boolean switch of loft features
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/boolean-switch-of-loft-features/td-p/13787202
' Category: advanced
' Scraped: 2025-10-07T13:08:59.922979

Dim oDoc As PartDocument = ThisDoc.Document 
 Dim oLofts As LoftFeatures = oDoc.ComponentDefinition.Features.LoftFeatures
 For Each oLoft As LoftFeature In oLofts
 
 If oLoft.Operation = PartFeatureOperationEnum.kJoinOperation Then 
		
	Dim LoftDefinition As LoftDefinition = oLoft.Definition
	
	LoftDefinition.Operation = PartFeatureOperationEnum.kCutOperation
	
	 oLoft.EnableAssociation = True 
	 oLoft.Suppressed = False
	End If

Next