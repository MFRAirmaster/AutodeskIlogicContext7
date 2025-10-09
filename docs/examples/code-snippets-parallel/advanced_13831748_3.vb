' Title: iLogic Rule to Measure Area of Two Faces
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-to-measure-area-of-two-faces/td-p/13831748#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:06:16.137540

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