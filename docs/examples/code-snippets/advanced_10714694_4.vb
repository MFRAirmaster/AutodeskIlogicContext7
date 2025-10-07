' Title: Create Model States using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-model-states-using-ilogic/td-p/10714694
' Category: advanced
' Scraped: 2025-10-07T13:06:06.974609

Dim doc As AssemblyDocument = ThisDoc.Document
Dim occ As ComponentOccurrence = doc.ComponentDefinition.Occurrences.Item(1)

Dim facDoc As PartDocument = occ.Definition.Document
If (facDoc.ComponentDefinition.IsModelStateMember) Then
    facDoc = facDoc.ComponentDefinition.FactoryDocument
End If
Dim modelStates As ModelStates = facDoc.ComponentDefinition.ModelStates
Dim newModelStateName = "YourNameHere"
Dim modelState As ModelState = modelStates.Add()
modelState.Name = newModelStateName
modelState.Activate()

occ.ActiveModelState = newModelStateName