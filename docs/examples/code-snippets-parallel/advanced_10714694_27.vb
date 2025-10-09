' Title: Create Model States using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-model-states-using-ilogic/td-p/10714694
' Category: advanced
' Scraped: 2025-10-09T08:50:38.589198

Dim Doc As PartDocument = ThisDoc.Document
oPCD = Doc.ComponentDefinition
If oPCD.IsModelStateMember Then
	Doc = oPCD.FactoryDocument
End If
oMStates = oPCD.ModelStates
newModelStateName = HEIGHT
oMState = oMStates.Add(newModelStateName)