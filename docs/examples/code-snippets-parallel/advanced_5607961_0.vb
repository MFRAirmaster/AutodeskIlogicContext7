' Title: Easy way to align view/scale label?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/easy-way-to-align-view-scale-label/td-p/5607961
' Category: advanced
' Scraped: 2025-10-07T14:00:38.554298

'27.11.2019
'No frills code to align the views along the bottom edge.

Sub Main()
Dim oSSet As SelectSet = ThisDoc.Document.SelectSet
Dim oItems As ObjectsEnumerator = ThisApplication.TransientObjects.CreateObjectCollection
Dim oDoc As Document = ThisDoc.Document
Dim oSelection As SelectSet = oDoc.SelectSet
'If oSelection Is Nothing Then
If oSSet.Count <2 Then
	MessageBox.Show("Nothing selected", "iLogic", MessageBoxButtons.OK)
	Exit Sub
End If
' Dim oItems As ObjectsEnumerator = ThisApplication.TransientObjects.CreateObjectCollection ' Removed for 2.1 as it is now declared earlier
For Each Item As Object In oSelection
	If Not TypeOf(Item) Is DrawingView Then Continue For 'DrawingViews only
	oItems.Add(Item)
Next

oDoc.SelectSet.SelectMultiple(oItems)
oHor_Bott(oItems)
End Sub

'=========================================================================================================
Sub oHor_Bott(oItems)
Dim oDrawDoc As DrawingDocument
oDrawDoc = ThisDoc.Document
Dim oView As DrawingView
Dim oSheet As Sheet
oSheet = oDrawDoc.ActiveSheet
Dim trans As Transaction
trans = ThisApplication.TransactionManager.StartTransaction(oDrawDoc, "Align Views")
'Dim oSSet As SelectSet = oDrawDoc.SelectSet
'For Each oView In oSSet
		Dim oFirstView As DrawingView = TryCast(oItems.item(1), DrawingView)
		Dim YPos As Double
		YPos = oFirstView.Position.Y - (oFirstView.Height/2)
        Dim oPoint2D As Inventor.Point2d
 
		 For Each oView In oItems
            oPoint2D = ThisApplication.TransientGeometry.CreatePoint2d(oView.Position.X,YPos + (oView.Height/2))
           oView.Position = oPoint2D
        Next
trans.End
Dim oDoc As Document = ThisDoc.Document
oDoc.SelectSet.SelectMultiple(oItems)
End Sub