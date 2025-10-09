' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454
' Category: advanced
' Scraped: 2025-10-09T09:00:46.753538

Sub Main
	Dim oBaseComponent As ComponentOccurrence = PickComponent("Pick Base Component.")
	oDoc = ThisDoc.Document
	If IsNothing(oBaseComponent) Then Exit Sub
		oBaseComponent.Grounded = True
	Dim oADef As AssemblyComponentDefinition = oBaseComponent.Parent
	Dim oConsts As AssemblyConstraints = oADef.Constraints
	oTrans = ThisApplication.TransactionManager.StartTransaction(oADef.Document, "Constrain Components (API)")
	Dim oOcc As ComponentOccurrence
	Dim oOccCol As Inventor.ObjectCollection
	oOccCol = ThisApplication.TransientObjects.CreateObjectCollection
	Dim oBaseWPPs As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(oBaseComponent)
	If IsNothing(oBaseWPPs) OrElse (Not oBaseWPPs.Count = 3) Then
		MsgBox("Failed to get 3 WorkPlaneProxy objects from Base component. Exiting.", vbCritical, "")
		oTrans.Abort
		Exit Sub
	End If
Do
	'Ask user for a selection
	oOcc = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select a component :")
	If Not oOcc Is  Nothing Then
	'Add component to collection
	oOccCol.Add(oOcc)
	'Set selected component inactive
	oOcc.Enabled = False
	End If
	Loop While Not oOcc Is Nothing 
	'Add collection to selection
	oDoc.SelectSet.SelectMultiple(oOccCol)
	oSelSet = oDoc.SelectSet
	
	'<<<< Good Point For Loop Of Remaining Code >>>>
	

		
For Each oOcc In oSelSet
		oOcc.Grounded = True
	Dim oOccWPPs As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(oOcc)
	If IsNothing(oOccWPPs) OrElse (Not oOccWPPs.Count = 3) Then
		MsgBox("Failed to get 3 WorkPlaneProxy objects from component to move. Exiting.", vbCritical, "")
		oTrans.Abort
		Exit Sub
	End If
'	DeleteConstraints(oOcc)

	

	For i As Integer = 0 To 2
		For i2 As Integer = 0 To 2
		Dim oParallel As Boolean = oBaseWPPs.Item(i).Plane.IsParallelTo(oOccWPPs.Item(i2).Plane, 0.00001)
		Dim oSameDirection As Boolean=False
		oSameDirection= oBaseWPPs.Item(i).Plane.Normal.IsEqualTo(oOccWPPs.Item(i2).Plane.Normal)
		oOffSet = oBaseWPPs.Item(i).Plane.DistanceTo(oOccWPPs.Item(i2).Plane.RootPoint)
'		msgbox(oOffset)
		If oParallel=True Then


				If oSameDirection = True Then
					Try
				oConstraint = oConsts.AddFlushConstraint(oBaseWPPs.Item(i), oOccWPPs.Item(i2), oOffSet)
				Catch
					End Try
				End If
				
				If oSameDirection = False Then
					Try
				oConstraint = oConsts.AddMateConstraint(oBaseWPPs.Item(i), oOccWPPs.Item(i2), oOffSet)
				Catch
					end try
				End If
	  			End If

		Next
	Next
	'<<<< End Loop Here, If Using One >>>>

	oOcc.Grounded = False
	oOcc.Enabled = True
	Next
	oBaseComponent.Grounded = False
	oDoc.SelectSet.Clear
	
	oTrans.End
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