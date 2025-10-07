' Title: Can we access textbox objects on the sheet object?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/can-we-access-textbox-objects-on-the-sheet-object/td-p/13795397#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:58:52.774033

Sub main

	Dim oDoc As DrawingDocument = ThisApplication.ActiveDocument
	Dim OTG As TransientGeometry = ThisApplication.TransientGeometry
	
	Dim oSheet As Sheet = oDoc.Sheets(4)
	
	Dim InputText As String = "Testing textboxes"
	Dim TextPoint As Point2d = OTG.CreatePoint2d(5, 5)
		
	Dim iNote As DrawingNote = oSheet.DrawingNotes.GeneralNotes.AddFitted(TextPoint, InputText)
	
	 



	
End Sub