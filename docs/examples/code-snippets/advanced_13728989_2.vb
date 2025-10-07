' Title: Using Named Entities to Automate Dimension using Inventor API.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/using-named-entities-to-automate-dimension-using-inventor-api/td-p/13728989
' Category: advanced
' Scraped: 2025-10-07T13:37:51.226505

Dim doc = ThisApplication.ActiveDocument
Dim osheet = doc.ActiveSheet
Dim oview = osheet.DrawingViews.Item(1)

' Get the referenced document from the drawing view
Dim refDoc
Try
	refDoc = oview.ReferencedDocumentDescriptor.ReferencedDocument
Catch ex As Exception
	MessageBox.Show("Unable to get referenced document: " & ex.Message)
	Return
End Try

If refDoc Is Nothing Then
	MessageBox.Show("Referenced document not available.")
	Return
End If

' Get named entities from the referenced document
Dim namedEntities
Try
	namedEntities = iLogicVb.Automation.GetNamedEntities(refDoc)
Catch ex As Exception
	MessageBox.Show("Failed to get named entities: " & ex.Message)
	Return
End Try

' Find named entities
Dim ent1 = namedEntities.FindEntity("BottomEdge")
Dim ent2 = namedEntities.FindEntity("TopEdge")

If ent1 Is Nothing Or ent2 Is Nothing Then
	MessageBox.Show("Named entity 'BottomEdge' or 'TopEdge' not found.")
	Return
End If

' Create geometry intents for dimensioning
Dim wp1 = osheet.CreateGeometryIntent(ent1)
Dim wp2 = osheet.CreateGeometryIntent(ent2)

' Set position of dimension
Dim tg = ThisApplication.TransientGeometry
Dim dimpos = tg.CreatePoint2d(oview.Position.X / 2, oview.Position.Y + 1)

' Add linear dimension
osheet.DrawingDimensions.GeneralDimensions.AddLinear(dimpos, wp1, wp2)