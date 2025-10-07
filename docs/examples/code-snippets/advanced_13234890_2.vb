' Title: iLogic Remove all Event Triggers &amp; add new Event Trigger
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-remove-all-event-triggers-amp-add-new-event-trigger/td-p/13234890#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:08:43.330465

Sub Main

	Dim oDoc As Document = ThisDoc.Document

	Dim oEvents As Inventor.PropertySet

	Try
		oEvents = oDoc.PropertySets.Item("{2C540830-0723-455E-A8E2-891722EB4C3E}")
	Catch
		oEvents = oDoc.PropertySets.Add("iLogicEventsRules", "{2C540830-0723-455E-A8E2-891722EB4C3E}")
	End Try

	For Each oItem As [Property] In oEvents
		oItem.Delete
	Next
	
	oEvents.Add("file://StockNumber-Assembly_Copy", "BeforeDocSave", 700)

End Sub