' Title: Adding Surface Texture to FlatPattern View
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/adding-surface-texture-to-flatpattern-view/td-p/13339496
' Category: advanced
' Scraped: 2025-10-07T14:10:25.946099

Dim drw As DrawingDocument = ThisDoc.Document
Dim osheet As Sheet = drw.ActiveSheet
For Each oDv As DrawingView In osheet.DrawingViews
	If oDv.IsFlatPatternView Then
		drw.SelectSet.Clear
		drw.SelectSet.Select(oDv)
		If oDv.Camera.Eye.X > oDv.Camera.Eye.Z Or oDv.Camera.Eye.Y > oDv.Camera.Eye.Z Then ' if flat pattern view is side view
			Dim oiptdoc As PartDocument = oDv.ReferencedDocumentDescriptor.ReferencedDocument
			Dim ft As FlatPattern = oiptdoc.ComponentDefinition.flatpattern
			Dim oFT As Object = ft.TopFace
			Dim oFB As Object = ft.BottomFace
			Dim oDvCT As DrawingCurve = oDv.DrawingCurves(oFT)(1) ' get first drawing view curve from top flatpattern face
	 		Dim oMidPointT As Point2d = oDvCT.MidPoint
			Dim oDvCB As DrawingCurve = oDv.DrawingCurves(oFB)(1) ' get first drawing view curve from bottomflatpattern face
	 		Dim oMidPointB As Point2d = oDvCB.MidPoint
	    	' Set a reference to the TransientGeometry object.
	    	Dim oTG As TransientGeometry= ThisApplication.TransientGeometry
	    	Dim oLeaderPointsT As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
	    	' Create a few leader points.
			If Abs(oMidPointT.X - oDv.Left) >= oDv.Width Then
				Call oLeaderPointsT.Add(oTG.CreatePoint2d(oDv.Left+ oDv.Width, oMidPointT.Y))
				Else
					Call oLeaderPointsT.Add(oTG.CreatePoint2d(oDv.Left , oMidPointT.Y))
				End If
	    	
	     	' Create an intent and add to the leader points collection.
		    ' This is the geometry that the leader text will attach to.
		    Dim oGeometryIntentT As GeometryIntent = osheet.CreateGeometryIntent(oDvCT,oMidPointT)
		'Call oLeaderPointsT.Add(oTG.CreatePoint2d(oMidPointT.X,oDv.Left))
		    Call oLeaderPointsT.Add(oGeometryIntentT)
			
		  	Dim oLeaderPointsB As ObjectCollection = ThisApplication.TransientObjects.CreateObjectCollection
		    ' Create a few leader points.
			
		  If Abs(oMidPointB.X - oDv.Left) >= oDv.Width Then
				Call oLeaderPointsB.Add(oTG.CreatePoint2d(oDv.Left, oMidPointB.Y))
				Else
					Call oLeaderPointsB.Add(oTG.CreatePoint2d(oDv.Left + oDv.Width, oMidPointB.Y))
				End If
		     ' Create an intent and add to the leader points collection.
		    ' This is the geometry that the leader text will attach to.
		    Dim oGeometryIntentB As GeometryIntent = osheet.CreateGeometryIntent(oDvCB,oMidPointB)
		    Call oLeaderPointsB.Add(oGeometryIntentB)
		
		 	Dim oSymbolT As SurfaceTextureSymbol
		 	oSymbolT = osheet.SurfaceTextureSymbols.Add(oLeaderPointsT,kMaterialRemovalProhibitedSurfaceType)
			Dim oSymbolB As SurfaceTextureSymbol
		    oSymbolB = osheet.SurfaceTextureSymbols.Add(oLeaderPointsB,kMaterialRemovalProhibitedSurfaceType)

		End If
	End If
Next