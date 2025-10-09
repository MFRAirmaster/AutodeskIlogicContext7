' Title: Using Ilogic to update model states
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/using-ilogic-to-update-model-states/td-p/12844817
' Category: api
' Scraped: 2025-10-09T09:00:40.424556

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