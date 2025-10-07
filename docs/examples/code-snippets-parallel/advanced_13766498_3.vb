' Title: Trying to detect if the All  Workfeatures button is checked or not.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/trying-to-detect-if-the-all-workfeatures-button-is-checked-or/td-p/13766498#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:59:44.764055

Dim doc As Document = ThisDoc.Document
If doc.ObjectVisibility.AllWorkFeatures Then
	MessageBox.Show("Work Features Are Visible")
End If