' Title: Batch PDF ignoring one file
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/batch-pdf-ignoring-one-file/td-p/13756957
' Category: advanced
' Scraped: 2025-10-07T13:15:49.381514

Try
		oAsmDrawingDoc = ThisApplication.Documents.Open(ThisDoc.ChangeExtension(".idw"), False)
	Catch
		MessageBox.Show("No top level assembly drawing found.", "Missing Drawing", MessageBoxButtons.OK)				
	End Try