' Title: Automating Update Model in Drawings
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automating-update-model-in-drawings/td-p/13783094#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:08:00.818851

Dim oDDoc As DrawingDocument = TryCast(ThisDoc.Document, DrawingDocument)
If oDDoc Is Nothing Then Return
Dim oMDoc As Inventor.Document = ThisDoc.ModelDocument
If oMDoc Is Nothing Then Return
'oMDoc.Rebuild2(True)
oMDoc.Update2(True)
oDDoc.Update2(True)
For Each oSheet As Inventor.Sheet In oDDoc.Sheets
	For Each oDView As DrawingView In oSheet.DrawingViews
		Dim oRFD As FileDescriptor = oDView.ReferencedDocumentDescriptor.ReferencedFileDescriptor
		oRFD.ReplaceReference(oRFD.FullFileName)
	Next 'oDView
	oSheet.Update()
Next 'oSheet