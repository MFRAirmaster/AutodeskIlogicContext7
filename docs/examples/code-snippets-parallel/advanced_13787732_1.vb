' Title: Inventor 2025 &quot;The file was modified by a rule that ran on check-in.&quot;
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/inventor-2025-quot-the-file-was-modified-by-a-rule-that-ran-on/td-p/13787732
' Category: advanced
' Scraped: 2025-10-09T08:51:49.081005

Try
	Logger.Info(ThisApplication.CommandManager.ActiveCommand)
Catch
	Logger.Info("No active command")
	Exit Sub
End Try

If ThisApplication.CommandManager.ActiveCommand = "VaultCheckin" Then
	Dim doc As DrawingDocument = ThisDoc.Document
	Dim r As New Random
	Dim pt As Point2d = ThisApplication.TransientGeometry.CreatePoint2d(8*2.54*r.NextDouble(), 11*2.54*r.NextDouble())
	doc.Sheets.Item(1).DrawingNotes.GeneralNotes.AddFitted(pt, "Test")
	ThisDoc.Save
End If