' Title: Easy way to align view/scale label?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/easy-way-to-align-view-scale-label/td-p/5607961#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:08:28.497227

'27.11.2019
'No frills code to align the views along the bottom selected Line.
'This is just for a quicker method, as this is the most used command.


Sub Main()
Dim oSSet As SelectSet = ThisDoc.Document.SelectSet
Dim oItems As ObjectsEnumerator = ThisApplication.TransientObjects.CreateObjectCollection
Dim oDoc As Document = ThisDoc.Document
Dim oSelection As SelectSet = oDoc.SelectSet
'If oSelection Is Nothing Then
If oSSet.Count <1 Then
	MessageBox.Show("Nothing selected", "iLogic", MessageBoxButtons.OK)
	Exit Sub
End If
For Each Item As Object In oSelection
	If Not TypeOf(Item) Is DrawingView Then Continue For 'DrawingViews only
	oItems.Add(Item)
Next
oDoc.SelectSet.SelectMultiple(oItems)
Align_Labels(oItems)
End Sub

'=========================================================================================================
Sub Align_Labels(oItems)
Dim oDrawDoc As DrawingDocument
oDrawDoc = ThisDoc.Document
Dim oView As DrawingView
Dim oSheet As Sheet
oSheet = oDrawDoc.ActiveSheet
Dim trans As Transaction
trans = ThisApplication.TransactionManager.StartTransaction(oDrawDoc, "Align Views")
'Dim oSSet As SelectSet = oDrawDoc.SelectSet
'
		Dim oNewPosition As Point2d
		oNewPOS = InputBox("Add Offset?", "Title", 1)
		oNewPOS /= 10
 
For Each oView In oItems
		  oNewPosition = ThisApplication.TransientGeometry.CreatePoint2d(oView.Center.X, oView.Center.Y - (oView.Height / 2)-( 1+oNewPOS))
		   oView.Label.Position = oNewPosition
		Next
        
trans.End
Dim oDoc As Document = ThisDoc.Document
oDoc.SelectSet.SelectMultiple(oItems)
End Sub