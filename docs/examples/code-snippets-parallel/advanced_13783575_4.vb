' Title: How to Check if Model State Exists Before Setting It in iLogic?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-check-if-model-state-exists-before-setting-it-in-ilogic/td-p/13783575#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:53:22.386621

MSName = "4-5-1548-05"
oDrwView = ActiveSheet.View("VIEW10")
Dim oModelStates As ModelStates = oDrwView.ModelFactoryDocument.ComponentDefinition.ModelStates
Dim MSfound As Boolean = False
For Each oModelState As ModelState In oModelStates
	If oModelState.Name.ToUpper = MSName.ToUpper Then
		MSfound = True
		Exit For
	End If
Next
If MSfound Then
	oDrwView.NativeEntity.SetActiveModelState(MSName)
Else
	MsgBox("Modelstate not found: " & MSName,,"Error ModelState")
End If