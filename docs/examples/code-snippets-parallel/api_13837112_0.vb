' Title: Invisible / abandoned weld seam symbols in drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/invisible-abandoned-weld-seam-symbols-in-drawing/td-p/13837112
' Category: api
' Scraped: 2025-10-09T09:05:53.817751

Dim oActiveDoc As DrawingDocument = ThisApplication.ActiveDocument

For Each oSheet As Sheet In oActiveDoc.Sheets

	For Each oDrawingWeldsymbol As DrawingWeldingSymbol In oSheet.WeldingSymbols

			MsgBox(oDrawingWeldsymbol.Position.X & vbCrLf & oDrawingWeldsymbol.Position.Y)
			
	Next

Next