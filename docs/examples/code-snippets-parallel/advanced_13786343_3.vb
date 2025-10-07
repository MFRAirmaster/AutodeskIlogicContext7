' Title: LEADER NOTE FILTER SETTINGS
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/leader-note-filter-settings/td-p/13786343
' Category: advanced
' Scraped: 2025-10-07T14:05:45.215964

Sub Main
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oDDoc As Inventor.DrawingDocument = TryCast(oInvApp.ActiveDocument, Inventor.DrawingDocument)
	If oDDoc Is Nothing Then
		MsgBox("A Drawing must be 'active' for this rule to work!", vbCritical, "Wrong Document Type")
		Return
	End If
	Dim oSheet As Inventor.Sheet = oDDoc.ActiveSheet
	Dim oDView As Inventor.DrawingView = oSheet.DrawingViews.Item(1)
	'get the model document being detailed within this drawing's views
	Dim oModelDoc As Inventor.Document = Nothing
	Try
		oModelDoc = oDView.ReferencedDocumentDescriptor.ReferencedDocument
	Catch
	End Try
	If oModelDoc Is Nothing Then
		MsgBox("No Model Document Found In First View On Active Sheet!", , "")
		Return
	End If
	Dim oDCurve As Inventor.DrawingCurve = oDView.DrawingCurves.Item(1)
	'this defines the 'connection' between a piece of geometry an an annotation
	Dim oGIntent As Inventor.GeometryIntent = oSheet.CreateGeometryIntent(oDCurve)
	'specify where the text of the leader should be (2 units outside top left corner of view)
	Dim oTxtPt As Inventor.Point2d = oInvApp.TransientGeometry.CreatePoint2d()
	oTxtPt.X = oDView.Left - 2
	oTxtPt.Y = oDView.Top + 2
	'put collection of leader node locations into a collection
	Dim oLPts As Inventor.ObjectCollection = oInvApp.TransientObjects.CreateObjectCollection()
	oLPts.Add(oGIntent)
	oLPts.Add(oTxtPt)
	Dim oLNotes As Inventor.LeaderNotes = oSheet.DrawingNotes.LeaderNotes
	'in order to specify which iProperty we want the note to be pointing to,
	'we need to be able to specify the 'internal name' of the PropertySet containing that Property
	'in your case, that will be the 'custom' PropertySet, which is always the fourth set
	Dim oCustomPropSet As Inventor.PropertySet = oModelDoc.PropertySets.Item(4)
	Dim sPropSetID As String = oCustomPropSet.InternalName
	'now try to get the custom Property
	Dim oCustomProp As Inventor.Property = Nothing
	Try
		oCustomProp = oCustomPropSet.Item("TAGG")
	Catch
	End Try
	If oCustomProp Is Nothing Then
		MsgBox("No custom Property named 'TAGG' found in model document!", , "")
		Return
	End If
	'get the ID of the custom Property, if it was found
	Dim sPropID As String = oCustomProp.PropId.ToString()
	'now define the FormattedText contents, incorporating those two pieces of information
	Dim sFText As String = "<Property Document='model' FormatID='" & sPropSetID & "' PropertyID='" & sPropID & "' />"
	'now create the LeaderNote, and specify that FormattedText
	Dim oLNote As Inventor.LeaderNote = oLNotes.Add(oLPts, sFText)
End Sub