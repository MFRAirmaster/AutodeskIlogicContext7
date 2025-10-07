' Title: iLogic Drawing Layer Visibility Toggle
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-drawing-layer-visibility-toggle/td-p/13816579#messageview_0
' Category: api
' Scraped: 2025-10-07T13:57:06.588368

If ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True
	ElseIf ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = True Then
	ThisDrawing.Document.StylesManager.Layers("Visible Narrow (ANSI)").Visible = False
	
End If

iLogicVb.UpdateWhenDone = True