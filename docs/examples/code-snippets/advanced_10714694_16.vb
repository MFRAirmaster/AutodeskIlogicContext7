' Title: Create Model States using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-model-states-using-ilogic/td-p/10714694
' Category: advanced
' Scraped: 2025-10-07T13:06:06.974609

'just using a dummy variable here,
'to preserve automation/triggering behavior of unquoted local parameter
oDV = ManwayType

'get the 'local' part document object (the document this rule is saved within)
Dim oPDoc As PartDocument = ThisDoc.Document
oPDef = oPDoc.ComponentDefinition

'if the parameter is numerical, and not unit-less,
'this will return 'database units' version of the parameter's Value,
'not the 'document units' version of it
'oManwayType = oPDef.Parameters.Item("ManwayType").Value

oDocName = ThisDoc.FileName(True)

'these two will get the 'document units' version of the parameter's Value
'oManwayType = ManwayType 'unquoted parameter name
oManwayType = Parameter(oDocName, "ManwayType")

oFeats = oPDef.Features

If oManwayType = 600 Then
	oFeats.Item("MWAY600 Centre").Suppressed = False
	oFeats.Item("MWAY600 Holes").Suppressed = False
	oFeats.Item("MWAY600 PCD").Suppressed = False
	oFeats.Item("MWAY800 Centre").Suppressed = True
	oFeats.Item("MWAY800 Holes").Suppressed = True
	oFeats.Item("MWAY800 PCD").Suppressed = True
	oFeats.Item("MWAY900 Centre").Suppressed = True
	oFeats.Item("MWAY900 Holes").Suppressed = True
	oFeats.Item("MWAY800 PCD").Suppressed = True
ElseIf oManwayType = 800 Then
	oFeats.Item("MWAY800 Centre").Suppressed = False
	oFeats.Item("MWAY800 Holes").Suppressed = False
	oFeats.Item("MWAY800 PCD").Suppressed = False
	oFeats.Item("MWAY600 Centre").Suppressed = True
	oFeats.Item("MWAY600 Holes").Suppressed = True
	oFeats.Item("MWAY600 PCD").Suppressed = True
	oFeats.Item("MWAY900 Centre").Suppressed = True
	oFeats.Item("MWAY900 Holes").Suppressed = True
	oFeats.Item("MWAY900 PCD").Suppressed = True
ElseIf oManwayType = 900 Then
	oFeats.Item("MWAY900 Centre").Suppressed = False
	oFeats.Item("MWAY900 Holes").Suppressed = False
	oFeats.Item("MWAY900 PCD").Suppressed = False
	oFeats.Item("MWAY800 Centre").Suppressed = True
	oFeats.Item("MWAY800 Holes").Suppressed = True
	oFeats.Item("MWAY800 PCD").Suppressed = True
	oFeats.Item("MWAY600 Centre").Suppressed = True
	oFeats.Item("MWAY600 Holes").Suppressed = True
	oFeats.Item("MWAY600 PCD").Suppressed = True
End If

oPDoc.Update