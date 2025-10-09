' Title: Batch PDF ignoring one file
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/batch-pdf-ignoring-one-file/td-p/13756957#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:58:52.070032

Try
		oAsmDrawingDoc = ThisApplication.Documents.Open(ThisDoc.ChangeExtension(".idw"), False)
	Catch
		MessageBox.Show("No top level assembly drawing found.", "Missing Drawing", MessageBoxButtons.OK)				
	End Try