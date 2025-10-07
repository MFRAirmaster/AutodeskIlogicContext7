' Title: iLogic get preparation drawing view document
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-get-preparation-drawing-view-document/td-p/13804570#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:16:19.862615

Dim oInvApp As Inventor.Application = ThisApplication
Dim oView As DrawingView
oView = oInvApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Select View...")
If oView Is Nothing Then Exit Sub
Dim eWeld As WeldmentStateEnum
Dim oObj As Object
Call oView.GetWeldmentState(eWeld, oObj)
If Not eWeld = WeldmentStateEnum.kPreparationsWeldmentState Then Exit Sub
Dim oOcc As ComponentOccurrence = oObj
MessageBox.Show(oOcc.Name, eWeld.ToString())