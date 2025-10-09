' Title: Ataching leader note to sketched symbol in drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ataching-leader-note-to-sketched-symbol-in-drawing/td-p/13833109#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:56:47.455439

Dim oDrawDoc As DrawingDocument
oDrawDoc = ThisApplication.ActiveDocument

Dim oSegment As DrawingCurveSegment
oSegment = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kDrawingCurveSegmentFilter, "Select an edge")

If oSegment is nothing then exit sub

Dim oSymbol As SketchedSymbol

Dim oActiveSheet As Sheet
oActiveSheet = oDrawDoc.ActiveSheet

oSymbol = oActiveSheet.SketchedSymbols(1)

Dim oMidPoint As Point2d = oSymbol.Position

Dim SelectPt As Point2d = oSegment.Parent.MidPoint

Dim oLeaderPoints As ObjectCollection
oLeaderPoints = ThisApplication.TransientObjects.CreateObjectCollection

Dim oGeometryIntent As GeometryIntent
oLeaderPoints.Add(oSymbol.Position)

oGeometryIntent = oActiveSheet.CreateGeometryIntent(oSegment.Parent)

Call oLeaderPoints.Add(oGeometryIntent)


'oSymbol.LeaderClipping = True
'oSymbol.LeaderVisible = True

Try
	oSymbol.Leader.AddLeader(oLeaderPoints)
Catch 
	'if the symbol already has a leader
	oSymbol.Leader.RootNode.AddLeader(oLeaderPoints)
End Try