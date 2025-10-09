' Title: iLogic Remove all Event Triggers &amp; add new Event Trigger
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-remove-all-event-triggers-amp-add-new-event-trigger/td-p/13234890#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:49:49.376821

Dim oDoc As Inventor.Document = ThisDoc.Document
Dim oSet As Inventor.PropertySet = Nothing
If oDoc.PropertySets.PropertySetExists("{2C540830-0723-455E-A8E2-891722EB4C3E}", oSet) Then
	Try : oSet.Delete() : Catch : End Try
End If