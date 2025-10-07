' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:33:58.654638

Sub Main
	Dim oBaseComponent As ComponentOccurrence = PickComponent("Pick Base Component.")
	If IsNothing(oBaseComponent) Then Exit Sub
		oBaseComponent.Grounded = True
	Dim oADef As AssemblyComponentDefinition = oBaseComponent.Parent
	Dim oConsts As AssemblyConstraints = oADef.Constraints
	oTrans = ThisApplication.TransactionManager.StartTransaction(oADef.Document, "Constrain Components (API)")
	Dim oBaseWPPs As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(oBaseComponent)
	If IsNothing(oBaseWPPs) OrElse (Not oBaseWPPs.Count = 3) Then
		MsgBox("Failed to get 3 WorkPlaneProxy objects from Base component. Exiting.", vbCritical, "")
		oTrans.Abort
		Exit Sub
	End If
	Dim oBaseTrans As Matrix = oBaseComponent.Transformation
	
	'<<<< Good Point For Loop Of Remaining Code >>>>
	Dim oCompToMove As ComponentOccurrence = PickComponent("Pick Component To Move.")
	If IsNothing(oCompToMove) Then oTrans.Abort : Exit Sub
		oCompToMove.Grounded = True
	Dim oCompToMoveWPPs As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(oCompToMove)
	If IsNothing(oCompToMoveWPPs) OrElse (Not oCompToMoveWPPs.Count = 3) Then
		MsgBox("Failed to get 3 WorkPlaneProxy objects from component to move. Exiting.", vbCritical, "")
		oTrans.Abort
		Exit Sub
	End If
	DeleteConstraints(oCompToMove)

	

	For i As Integer = 0 To 2
		For i2 As Integer = 0 To 2
		Dim oParallel As Boolean = oBaseWPPs.Item(i).Plane.IsParallelTo(oCompToMoveWPPs.Item(i2).Plane, 0.00001)
		Dim oSameDirection As Boolean=False
		oSameDirection= oBaseWPPs.Item(i).Plane.Normal.IsEqualTo(oCompToMoveWPPs.Item(i2).Plane.Normal)
		oOffSet = oBaseWPPs.Item(i).Plane.DistanceTo(oCompToMoveWPPs.Item(i2).Plane.RootPoint)
		If oParallel=True Then
'

				If oSameDirection=True Then
				oConstraint=oConsts.AddFlushConstraint(oBaseWPPs.Item(i), oCompToMoveWPPs.Item(i2), oOffSet)
				End If
				If oSameDirection=False Then
				oConstraint=oConsts.AddMateConstraint(oBaseWPPs.Item(i), oCompToMoveWPPs.Item(i2), oOffSet)
				End If
	  			End If

		Next
	Next
	'<<<< End Loop Here, If Using One >>>>
	oTrans.End
	oCompToMove.Grounded = False
	oBaseComponent.Grounded = False
End Sub

Function PickComponent(oPrompt As String) As ComponentOccurrence
	oObj = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, oPrompt)
	If IsNothing(oObj) OrElse (TypeOf oObj Is ComponentOccurrence = False) Then Return Nothing
	Dim oOcc As ComponentOccurrence = oObj
	Return oOcc
End Function

Function GetComponentOriginPlaneProxies(oComp As ComponentOccurrence) As List(Of WorkPlaneProxy)
	If IsNothing(oComp) Then Return Nothing
	Dim oWPs As WorkPlanes = Nothing
	Dim oWPPs As New List(Of WorkPlaneProxy)
	If TypeOf oComp.Definition Is PartComponentDefinition Or _
		TypeOf oComp.Definition Is AssemblyComponentDefinition Then
		oWPs = oComp.Definition.WorkPlanes
	Else
		Return Nothing
	End If
	For i As Integer = 1 To 3
		Dim oWPP As WorkPlaneProxy = Nothing
		oWP = oWPs.Item(i)
		oComp.CreateGeometryProxy(oWP, oWPP)
		oWPPs.Add(oWPP)
	Next
	Return oWPPs
End Function

Sub DeleteConstraints(oComp As ComponentOccurrence)
	If IsNothing(oComp) Then Exit Sub
	If oComp.Constraints.Count = 0 Then Exit Sub
	For Each oConst As AssemblyConstraint In oComp.Constraints
		oConst.Delete
	Next
End Sub