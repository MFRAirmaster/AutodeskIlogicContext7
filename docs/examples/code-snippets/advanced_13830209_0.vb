' Title: inest file giving wrong Document Type Enumerator
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/inest-file-giving-wrong-document-type-enumerator/td-p/13830209
' Category: advanced
' Scraped: 2025-10-07T12:54:20.665665

Logger.Debug("Open Documents: " & ThisApplication.Documents.Count)

Dim oUnsavedDocs As New ArrayList 

Logger.Debug("View Frames " & ThisApplication.ViewFrames.Count)
Logger.Debug("Views " & ThisApplication.Views.Count)

For Each oAppVeiw As Inventor.View In ThisApplication.Views
'For Each oDoc As Inventor.Document In ThisApplication.Documents
	oDoc = oAppVeiw.Document
	
	Logger.Debug(oDoc.DisplayName)
	Logger.Debug("Dirty: " & oDoc.Dirty)
	
	Logger.Debug("Type: " & oDoc.DocumentType.ToString)
	
	If oDoc.FileSaveCounter<1 Then
		oUnsavedDocs.Add(oDoc)
		Logger.Debug("Times Saved: " & oDoc.FileSaveCounter)
	End If 
	
	If oDoc.DocumentType = kNestingDocument Then
		Logger.Debug("HELLO")
	End If 
	
	Next
	Logger.Debug("Size of oUnsavedDocs: " & oUnsavedDocs.Count)