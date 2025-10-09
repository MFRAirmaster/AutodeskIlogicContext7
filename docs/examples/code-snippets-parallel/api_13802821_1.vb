' Title: Use filename of 2D file, to fill in iproperties - Inventor 2024
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/use-filename-of-2d-file-to-fill-in-iproperties-inventor-2024/td-p/13802821#messageview_0
' Category: api
' Scraped: 2025-10-09T09:08:18.866703

Try
	Dim oDoc As Document = ThisApplication.ActiveEditDocument
	Dim oName As String = ThisDoc.FileName(False)
	Dim oFirstDashPos As Integer = InStr(oName, "-")
	Dim oSecondDashPos As Integer = InStr(oFirstDashPos + 1, oName, "-")
	Dim oTitle As String = Left(oName, oFirstDashPos -1)
	Dim oClient As String = Mid(oName, oFirstDashPos + 1, oSecondDashPos - oFirstDashPos - 1)
	Dim oDesc As String = Right(oName, Len(oName) - oSecondDashPos)
	
	iProperties.Value("Summary", "Title") = oTitle
	iProperties.Value("Custom", "Client") = oClient
	iProperties.Value("Project", "Description") = oDesc
Catch
End Try