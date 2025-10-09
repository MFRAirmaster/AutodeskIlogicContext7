' Title: Create Model States using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-model-states-using-ilogic/td-p/10714694
' Category: advanced
' Scraped: 2025-10-09T08:50:38.589198

Dim oDoc As AssemblyDocument
	
Dim oOcc As ComponentOccurrence 	For Each oOcc In Occurrences    'The below line does not work, what is the correct syntax????    oOcc.ComponentDefinition.ModelStates.Add(oName)next