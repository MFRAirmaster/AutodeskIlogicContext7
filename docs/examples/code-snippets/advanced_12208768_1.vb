' Title: Promote a model state to become the master state with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/promote-a-model-state-to-become-the-master-state-with-ilogic/td-p/12208768
' Category: advanced
' Scraped: 2025-10-07T13:05:48.627090

Dim oDoc As Document = ThisDoc.FactoryDocument
If oDoc Is Nothing Then Exit Sub
Dim oMSs As ModelStates
Try : oMSs = oDoc.ComponentDefinition.ModelStates : Catch : End Try
If oMSs Is Nothing OrElse oMSs.Count < 2 Then Exit Sub
Dim oPrimaryMS As ModelState = oMSs.Item(1) 'always the first one
Dim oActiveMS As ModelState = oMSs.ActiveModelState
If oActiveMS Is oPrimaryMS Then Exit Sub 'or could simply delete all custom ModelStates
Dim oPrimaryRow As ModelStateTableRow = Nothing
Dim oActiveRow As ModelStateTableRow = Nothing
Dim oMSTable As ModelStateTable = oMSs.ModelStateTable
Dim oRows As ModelStateTableRows = oMSTable.TableRows
For Each oRow As ModelStateTableRow In oRows
	If oRow.MemberName = oPrimaryMS.Name Then
		oPrimaryRow = oRow
	ElseIf oRow.MemberName = oActiveMS.Name Then
		oActiveRow = oRow
	Else
		Try : oRow.Delete : Catch : End Try
	End If
Next 'oRow
For i As Integer = 2 To oActiveRow.Count
	'Logger.Info(oActiveRow.Item(i).Value)
	Try : oPrimaryRow.Item(i).Value = oActiveRow.Item(i).Value : Catch : End Try
Next
oPrimaryMS.Activate
Try : oActiveMS.Delete : Catch : End Try