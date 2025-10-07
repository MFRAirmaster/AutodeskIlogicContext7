' Title: ilogic - Edit rectangular pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-edit-rectangular-pattern/td-p/8740998
' Category: advanced
' Scraped: 2025-10-07T13:35:54.731250

Dim oDoc As AssemblyDocument
oDoc = ThisApplication.ActiveDocument

Dim oChannels As ObjectCollection
oChannels = ThisApplication.TransientObjects.CreateObjectCollection

If oChannels.Count > 0
	oChannels.Clear
End If

'add appropriate channels to object collection
Select Num_Splits
Case 0
	oChannels.Add("Channel HT")
Case 1
	oChannels.Add("Channel H1")
	oChannels.Add("Channel 1T")
Case 2
	oChannels.Add("Channel H1")
	oChannels.Add("Channel 12")
	oChannels.Add("Channel 2T")
Case 3
	oChannels.Add("Channel H1")
	oChannels.Add("Channel 12")
	oChannels.Add("Channel 23")
	oChannels.Add("Channel 3T")
End Select


Dim oPatt As RectangularOccurrencePattern 

'get pattern
oPatt = oDoc.ComponentDefinition.OccurrencePatterns.Item("Channel Pattern")

'change pattern components 
oPatt.ParentComponents = oChannels
 
InventorVb.DocumentUpdate()