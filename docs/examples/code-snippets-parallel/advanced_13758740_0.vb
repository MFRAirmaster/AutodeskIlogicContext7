' Title: Dimension Visibility
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/dimension-visibility/td-p/13758740
' Category: advanced
' Scraped: 2025-10-09T09:02:24.526297

If Parameter("STRAIGHT CONVEYOR.iam.Legs") = "No" Then
oHideOption = True
		oDoc = ThisDoc.Document

		Try
			'try to create the layer
			oOrphanLayer = oDoc.StylesManager.Layers.Item(1).Copy("Orphan Dims")
		Catch
			'assume error means layer already exists and
			'assign layer to variable
			oOrphanLayer = oDoc.StylesManager.Layers.Item("Orphan Dims")
			oLayer = oDoc.StylesManager.Layers.Item("Dimension(ISO)")
		End Try

		If oHideOption = True Then
			'Loop through all dimensions
			Dim oDrawingDim As DrawingDimension
			For Each oDrawingDim In oDoc.ActiveSheet.DrawingDimensions
				'look at only unattached dims
				If oDrawingDim.Attached = False Then
					' set the layer to the dummy layer 
					oDrawingDim.Layer = oOrphanLayer
				Else
					oDrawingDim.Layer = oLayer 
				End If
			Next
			'hide the layer 
			oOrphanLayer.Visible = False
		Else
			'show the layer 
			oOrphanLayer.Visible = True
		End If
		oDoc.Update
	End If