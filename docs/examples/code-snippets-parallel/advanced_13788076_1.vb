' Title: Sketch symbol &amp; Table creation for QC Inspection
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sketch-symbol-amp-table-creation-for-qc-inspection/td-p/13788076
' Category: advanced
' Scraped: 2025-10-09T09:01:52.259838

Dim oSkSybl As SketchedSymbol = ThisApplication.CommandManager.Pick( _
SelectionFilterEnum.kDrawingSketchedSymbolFilter, "Select a SketchedSymbol to inspect.")
If oSkSybl Is Nothing Then Return
Dim oLeader As Inventor.Leader = oSkSybl.Leader
If oLeader.HasRootNode Then
	Dim oRNode As LeaderNode = oLeader.RootNode
	Dim oNodes As LeaderNodesEnumerator = oLeader.AllNodes
	For Each oNode As LeaderNode In oNodes
		If oNode.AttachedEntity Is Nothing Then Continue For
		Dim oGI As Inventor.GeometryIntent = oNode.AttachedEntity
		Dim oGeom As Object = Nothing : Try : oGeom = oGI.Geometry : Catch : End Try
		If oGeom Is Nothing Then Continue For
		MsgBox("TypeName(oGeom) = " & TypeName(oGeom),,"")
		If TypeOf oGeom Is DrawingDimension Then
			Dim oDDim As DrawingDimension = oGeom
			MsgBox(oDDim.Text.Text,,"")
		End If
	Next 'oNode
End If