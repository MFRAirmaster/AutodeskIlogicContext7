' Title: Update All Model States of ipt.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-all-model-states-of-ipt/td-p/13824102
' Category: advanced
' Scraped: 2025-10-09T09:05:53.816586

Dim oDoc As Inventor.Document = ThisApplication.ActiveDocument
If oDoc Is Nothing Then Return
If (Not TypeOf oDoc Is AssemblyDocument) AndAlso (Not TypeOf oDoc Is PartDocument) Then Return
If oDoc.ComponentDefinition.IsModelStateMember Then
	oDoc = oDoc.ComponentDefinition.FactoryDocument
End If
Dim oMSs As ModelStates = oDoc.ComponentDefinition.ModelStates
If (oMSs Is Nothing) OrElse (oMSs.Count < 2) Then Return
Dim oAMS As ModelState = oMSs.ActiveModelState
For Each oMS As ModelState In oMSs
	oMS.Activate()
	Dim oFDoc As Inventor.Document = ThisApplication.ActiveDocument
	oFDoc.Update2(True)
Next oMS
If oMSs.ActiveModelState IsNot oAMS Then oAMS.Activate()