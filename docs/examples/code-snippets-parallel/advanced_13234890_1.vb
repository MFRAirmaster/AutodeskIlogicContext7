' Title: iLogic Remove all Event Triggers &amp; add new Event Trigger
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-remove-all-event-triggers-amp-add-new-event-trigger/td-p/13234890#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:06:50.006076

Sub Main()
	

RemoveRulesfromEventTrigger
AssemblyEvents


End Sub


Function GetiLogicEventPropSet(cDocument As Document) As Inventor.PropertySet
	On Error Resume Next
	iLogicEventPropSet = cDocument.PropertySets.Item("iLogicEventsRules")

	If iLogicEventPropSet Is Nothing Then
		iLogicEventPropSet = cDocument.PropertySets.Item("_iLogicEventsRules")
	End If

	If iLogicEventPropSet.InternalName <> "{2C540830-0723-455E-A8E2-891722EB4C3E}" Then
		Call iLogicEventPropSet.Delete
		iLogicEventPropSet = cDocument.PropertySets.Add("iLogicEventsRules", "{2C540830-0723-455E-A8E2-891722EB4C3E}")
	End If

	If iLogicEventPropSet Is Nothing Then
		iLogicEventPropSet = cDocument.PropertySets.Add("iLogicEventsRules", "{2C540830-0723-455E-A8E2-891722EB4C3E}")
	End If

	If iLogicEventPropSet Is Nothing Then
		MsgBox("Unable to create the Event Triggers property for this file!", , "Event Triggers Not Set")
		Err.Raise(1)
		Exit Function
	End If
	On Error GoTo 0

	Return iLogicEventPropSet
End Function


Sub RemoveRulesfromEventTrigger()


   Try
Dim EventTrigger = ThisDoc.Document.PropertySets.Item("{2C540830-0723-455E-A8E2-891722EB4C3E}")
'MsgBox(x.Count)

If EventTrigger.Name = "StockNumber-Assembly_Copy" Then
	Return
	Else
		ThisDoc.Document.PropertySets.Item("{2C540830-0723-455E-A8E2-891722EB4C3E}").Delete
	End If
	Catch
	End Try

End Sub

Sub AssemblyEvents



	On Error Resume Next
	Dim EventPropSet As Inventor.PropertySet
	EventPropSet = GetiLogicEventPropSet(ThisApplication.ActiveDocument)
	' To make sure that the document has an iLogic DocumentInterest, add a temporary rule
	'EventPropSet.Add("file://Accessory List", "AfterDrawingViewsUpdate", 1801)
	EventPropSet.Add("file://StockNumber-Assembly_Copy", "BeforeDocSave", 700)

	'After Open Document					: AfterDocOpen                 		: 400
	'Close(Document)						: DocClose                     		: 500
	'Before Save Document                   : BeforeDocSave           			: 700
	'After Save Document               		: AfterDocSave               		: 800
	'Any Model Parameter Change        		: AfterAnyParamChange   			: 1000
	'Part Geometry Change**            		: PartBodyChanged         			: 1200
	'Material Change**                  	: AfterMaterialChange     			: 1400
	'Drawing View Change***               	: AfterDrawingViewsUpdate  			: 1500
	'iProperty(Change)                  	: AfterAnyiPropertyChange           : 1600
	'Feature Suppression Change**          	: AfterFeatureSuppressionChange   	: 2000
	'Component Suppression Change*   		: AfterComponentSuppressionChange 	: 2200
	'iPart / iAssembly Change Component* 	: AfterComponentReplace   			: 2400
	'New Document                         	: AfterDocNew                  		: 2600

	InventorVb.DocumentUpdate()


End Sub