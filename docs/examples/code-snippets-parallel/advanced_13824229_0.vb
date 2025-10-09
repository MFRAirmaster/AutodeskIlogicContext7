' Title: Update all Model States at Once
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-all-model-states-at-once/td-p/13824229#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:23.487731

Dim oDoc As Inventor.Document = ThisDoc.FactoryDocument
If oDoc Is Nothing Then Return
Dim oMSs As ModelStates = oDoc.ComponentDefinition.ModelStates
If oMSs Is Nothing OrElse oMSs.Count = 0 Then Return
Dim oAMS As ModelState = oMSs.ActiveModelState
For Each oMS As ModelState In oMSs
	oMS.Activate
	Dim oMSDoc As Inventor.Document = TryCast(oMS.Document, Inventor.Document)
	If oMSDoc Is Nothing Then Continue For
	oMSDoc.Update2(True)
Next oMS
If oMSs.ActiveModelState IsNot oAMS Then oAMS.Activate