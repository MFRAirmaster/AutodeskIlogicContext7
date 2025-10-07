' Title: Create Model States using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-model-states-using-ilogic/td-p/10714694
' Category: advanced
' Scraped: 2025-10-07T13:06:06.974609

Dim doc As AssemblyDocument = ThisDoc.Document
oName = "COURSE 1 STEEL TANK SHEET:8"
occ = doc.ComponentDefinition.Occurrences.ItemByName(oName)
Dim facDoc As PartDocument = occ.Definition.Document
oPCD = facDoc.ComponentDefinition
If oPCD.IsModelStateMember Then
	facDoc = oPCD.FactoryDocument
End If
oMStates = oPCD.ModelStates
newModelStateName = "N1"
oMState = oMStates.Add(newModelStateName)
oMState.Activate()
occ.ActiveModelState = newModelStateName

oMState.FactoryDocument.ComponentDefinition.Features.Item("MWAY900 Centre").Suppressed = False
oMState.FactoryDocument.ComponentDefinition.Features.Item("MWAY900 Holes").Suppressed = False
oMState.FactoryDocument.ComponentDefinition.Features.Item("MWAY900 PCD").Suppressed = False