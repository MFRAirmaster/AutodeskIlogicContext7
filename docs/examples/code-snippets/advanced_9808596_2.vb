' Title: ilogic rule that creating a cope of a drawing and changing its iPart instance.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-that-creating-a-cope-of-a-drawing-and-changing-its/td-p/9808596
' Category: advanced
' Scraped: 2025-10-07T12:56:28.086085

Sub Main
	Dim oDDoc As DrawingDocument = TryCast(ThisApplication.ActiveDocument, Inventor.DrawingDocument)
	If oDDoc Is Nothing Then Return
	Dim oOrigSheet As Inventor.Sheet = oDDoc.ActiveSheet
	Dim oNewSheet As Inventor.Sheet = CopySheet(oDDoc, oOrigSheet)
End Sub

Function CopySheet(oDDoc As DrawingDocument, oSheet As Inventor.Sheet) As Inventor.Sheet
	Dim oSS As SelectSet = oDDoc.SelectSet
	oSS.Clear()
	Dim oCM As CommandManager = ThisApplication.CommandManager
	Dim oCDs As ControlDefinitions = oCM.ControlDefinitions
	Dim oCopy As ControlDefinition = oCDs.Item("AppCopyCmd")
	Dim oPaste As ControlDefinition = oCDs.Item("AppPasteCmd")
	oSS.Select(oSheet)
	oCopy.Execute2(True)
	oSS.Clear()
	oPaste.Execute2(True)
	Return oDDoc.Sheets.Item(oDDoc.Sheets.Count)
End Function