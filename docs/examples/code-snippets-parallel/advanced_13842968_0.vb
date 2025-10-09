' Title: Converting an Inventor .ipt model into Ghost / Reference Geometry using ClientGraphics.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/converting-an-inventor-ipt-model-into-ghost-reference-geometry/td-p/13842968#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:08:01.998946

Sub Main
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oPDoc As PartDocument = TryCast(oInvApp.ActiveDocument, Inventor.PartDocument)
	If oPDoc Is Nothing Then Return
	If oPDoc.RequiresUpdate Then oPDoc.Update2(True)
	'get first regular body in this part
	Dim oFirstSBody As SurfaceBody = oPDoc.ComponentDefinition.SurfaceBodies.Item(1)
	'copy it to a 'transient' body (not visible, just mathematical data)
	Dim oCopiedSBody As SurfaceBody = oInvApp.TransientBRep.Copy(oFirstSBody)
	'create a visible, but non-interactive, representation of the cloned body
	Dim oCG As ClientGraphics = oPDoc.NonTransactingClientGraphicsCollection.Add("MySurfaceBodyClone")
	oCG.Visible = True
	Dim oGNode As GraphicsNode = oCG.AddNode(1)
	oGNode.Visible = True
	Dim oSG As SurfaceGraphics = oGNode.AddSurfaceGraphics(oCopiedSBody)
	oInvApp.ActiveView.Update()
End Sub