' Title: How to Check if Model State Exists Before Setting It in iLogic?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-check-if-model-state-exists-before-setting-it-in-ilogic/td-p/13783575#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:53:22.386621

iLogicForm.Show("Form 1") '
ModStateName = DWG
ModStateName2 = DWG & "'"
oDrwView = ActiveSheet.View("VIEW15")
Dim oModelStates As ModelStates = oDrwView.ModelFactoryDocument.ComponentDefinition.ModelStates
Dim MSfound As Boolean = False
For Each oModelState As ModelState In oModelStates
	If oModelState.Name.ToUpper = ModStateName2.ToUpper Then
		MSfound = True
		Exit For
	End If
Next

If MSfound Then
	ActiveSheet.View("VIEW15").View.Suppressed = False
	ActiveSheet.View("VIEW15").NativeEntity.SetActiveModelState(ModStateName2)
Else
	ActiveSheet.View("VIEW15").View.Suppressed = True
	MsgBox("Modelstate not found: " & MSName,,"Error ModelState")
End If

ActiveSheet.View("VIEW10").NativeEntity.SetActiveModelState(ModStateName)