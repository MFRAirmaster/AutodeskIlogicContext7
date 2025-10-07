' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:33:58.654638

Sub Main
    Dim oBaseComponent As ComponentOccurrence = PickComponent("Pick Base Component.")
    If IsNothing(oBaseComponent) Then Exit Sub
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
    '''    Dim oAsmComp As AssemblyComponentDefinition= ThisApplication.ActiveDocument.ComponentDefinition
    For Each oOccurrence In oAsmComp.Occurrences
'    'ground everything in the top level
    oOccurrence.Grounded = True
    Next
    '''    '<<<< Good Point For Loop Of Remaining Code >>>>
'''    Dim oCompToMove As ComponentOccurrence = PickComponent("Pick Component To Move.")'''    If IsNothing(oCompToMove) Then oTrans.Abort : Exit Sub'''    Dim oCompToMoveWPPs As List(Of WorkPlaneProxy) = GetComponentOriginPlaneProxies(oCompToMove)'''    If IsNothing(oCompToMoveWPPs) OrElse (Not oCompToMoveWPPs.Count = 3) Then'''        MsgBox("Failed to get 3 WorkPlaneProxy objects from component to move. Exiting.", vbCritical, "")'''        oTrans.Abort'''        Exit Sub'''    End If

'''    DeleteConstraints(oCompToMove)'''    oCompToMove.Transformation = oBaseTrans'''    
'''    For i As Integer = 0 To 2'''        oConst = oConsts.AddFlushConstraint(oBaseWPPs.Item(i), oCompToMoveWPPs.Item(i), 0)'''    Next

    Dim oAsm As AssemblyDocument= ThisApplication.ActiveDocument
    Dim oUM As UnitsOfMeasure = oAsm.UnitsOfMeasure
    Dim oOcc As ComponentOccurrence
    For Each oOcc In oAsm.ComponentDefinition.Occurrences
        
        Dim oTransform As Matrix
        oTransform = oOcc.Transformation
        
        Dim oOriginLocation As Vector
        oOriginLocation = oTransform.Translation
        
        
        Dim AbstandvonEbene(0 To 3) As Double
        
        AbstandvonEbene(1) = oOriginLocation.X
        AbstandvonEbene(2) = oOriginLocation.Y
        AbstandvonEbene(3) = oOriginLocation.Z
        
        'Create a proxy for Face0 (The face in the context of the assembly)
        Dim Z�hler As Integer = 1
        
        'cycle each Origin plane in the top assembly
        For Z�hler = 1 To 3
        
            Dim oBaseWPPstOrPlane As WorkPlane = oBaseWPPs(Z�hler)    
            
            Dim oCompAsmDef As AssemblyComponentDefinition
            Dim oCompPtDef As PartComponentDefinition
            'cycle each Origin plane of base component in the first level Occurence in the assembly
            For Z�hler2 = 1 To 3
                
                If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    oCompAsmDef = oOcc.Definition
                    oOcc.CreateGeometryProxy(oCompAsmDef.WorkPlanes.Item(Z�hler2),curCompOriPlane)
                
                ElseIf oOcc.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    oCompPtDef = oOcc.Definition
                    oOcc.CreateGeometryProxy(oCompPtDef.WorkPlanes.Item(Z�hler2), curCompOriPlane)
                                
                End If
                
                Dim oParalell As Boolean= oBaseWPPstOrPlane.Plane.IsParallelTo(curCompOriPlane.Plane, 0.00001)
                Dim oSameDirection As Boolean= oBaseWPPstOrPlane.Plane.Normal.IsEqualTo(curCompOriPlane.Plane.Normal)
            
                Dim oNV As NameValueMap
            
                If oParalell=True Then
                    If oSameDirection=True Then
                    oConstraint=oAsmComp.Constraints.AddFlushConstraint(curAsmOrPlane, curCompOriPlane, AbstandvonEbene(Z�hler))
                    End If
                    If oSameDirection=False Then
                    oConstraint=oAsmComp.Constraints.AddMateConstraint(curAsmOrPlane, curCompOriPlane, AbstandvonEbene(Z�hler))
                    End If
            
                End If
            Next
        Next    
Next
'''    '<<<< End Loop Here, If Using One >>>>
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