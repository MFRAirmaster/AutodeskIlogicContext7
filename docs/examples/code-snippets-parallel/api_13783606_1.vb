' Title: Macro to delete specific sketch symbol?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/macro-to-delete-specific-sketch-symbol/td-p/13783606#messageview_0
' Category: api
' Scraped: 2025-10-07T14:00:17.398424

Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, Inventor.DrawingDocument)
If oDDoc Is Nothing Then Return
For Each oSheet As Inventor.Sheet In oDDoc.Sheets
	For Each oSS As SketchedSymbol In oSheet.SketchedSymbols
		If oSS.Name = "rev triangle" Then
			oSS.Delete()
		End If
	Next 'oSS
Next 'oSheet