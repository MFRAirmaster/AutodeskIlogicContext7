' Title: Cut a solid by RevolveFeature
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/cut-a-solid-by-revolvefeature/td-p/13837083#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:08:28.488479

Dim BackTrimCol As ObjectCollection = oTO.CreateObjectCollection
			
			Dim BackTrimSketch As PlanarSketch 
			
			Dim BackProfile As Profile
			
			Dim BackStartPoint As SketchPoint, BackRectPoint As SketchPoint, Rect2StartPoint As SketchPoint, Rect2EndPoint As SketchPoint
			
			Dim Rectangle2 As SketchEntitiesEnumerator
			
			Dim Rect2Line1 As SketchLine, Rect2Line3 As SketchLine
			
			Dim BackStartDis As DimensionConstraint, Rect2Line1Con As DimensionConstraint, CenterControlA As DimensionConstraint, Rect2Line3con As DimensionConstraint, CenterControlB As DimensionConstraint
			
			Dim BackRevolveTrim as RevolveFeature
				
					
					BackTrimSketch = oCompDef.Sketches.Add(oUCS.YZPlane)
						BackTrimSketch.Name = FittingNameInput & " upper trim sketch "
					
					BackStartPoint = BackTrimSketch.AddByProjectingEntity(Proj_Point)
					
					BackRectPoint  = BackTrimSketch.SketchPoints.Add(oTG.CreatePoint2d(BackStartPoint.Geometry.X, BackStartPoint.Geometry.Y - 1)) 
						BackTrimSketch.GeometricConstraints.AddVerticalAlign(BackStartPoint, BackRectPoint)
						
						BackStartDis = BackTrimSketch.DimensionConstraints.AddTwoPointDistance(BackRectPoint, BackStartPoint, Inventor.DimensionOrientationEnum.kAlignedDim, oTG.CreatePoint2d(10, 10))
							BackStartDis.Parameter.Expression = FittingNameInput & "_Fitting_Proj" & "*2 " & "+tank_id+sidewall_th*2"
							
						oDoc.Update	
							
						Rect2StartPoint = BackTrimSketch.SketchPoints.Add(oTG.CreatePoint2d(BackRectPoint.Geometry.X - 5, BackRectPoint.Geometry.Y - 5))
						Rect2EndPoint = BackTrimSketch.SketchPoints.Add(oTG.CreatePoint2d(BackRectPoint.Geometry.X + 5, BackRectPoint.Geometry.Y + 5))
							
						Rectangle2 = BackTrimSketch.SketchLines.AddAsTwoPointRectangle(Rect2StartPoint, Rect2EndPoint)	
							
							Rect2Line1 = Rectangle2.Item(2)
							Rect2Line3 = Rectangle2.Item(3)
							
							Rect2Line1Con = BackTrimSketch.DimensionConstraints.AddTwoPointDistance(Rect2Line1.StartSketchPoint, Rect2Line1.EndSketchPoint, Inventor.DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(3, 3))
								Rect2Line1Con.Parameter.Expression = "(6+sidewall_th+3)*2+12"
									
							CenterControlA = BackTrimSketch.DimensionConstraints.AddTwoPointDistance(BackRectPoint, Rect2EndPoint, Inventor.DimensionOrientationEnum.kVerticalDim, oTG.CreatePoint2d(-6, -6))
								CenterControlA.Parameter.Expression = "(6+sidewall_th)+3"
								
							Rect2Line3con= BackTrimSketch.DimensionConstraints.AddTwoPointDistance(Rect2Line3.StartSketchPoint, Rect2Line3.EndSketchPoint, Inventor.DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(-3, -3))
								Rect2Line3Con.Parameter.Value = 48 * 2.54
								
							CenterControlB = BackTrimSketch.DimensionConstraints.AddTwoPointDistance(BackRectPoint, Rect2StartPoint, Inventor.DimensionOrientationEnum.kHorizontalDim, oTG.CreatePoint2d(6, 6))
								CenterControlB.Parameter.Value = 24 * 2.54 
									
						BackProfile = BackTrimSketch.Profiles.AddForSolid
						
							oDoc.Update
						
								'The below is calculating the area to cut away from the outside
								Dim Ratio As Double =  (oUserParams.Item(FittingNameInput & "_Flange_Size").Value / 2) / (tank_id / 2) 
								Dim ThetaRad As Double = 2 * Math.Asin(Ratio)
								Dim ThetaDeg As Double = ThetaRad * (180 / Math.PI)
								
								Dim AdjustedAngle As Double = 360 - ThetaDeg 
								Dim AdjustedRad As Double = AdjustedAngle * Math.PI /180
								
					
						BackRevolveTrim = oCompDef.Features.RevolveFeatures.AddByAngle(BackProfile, oUCS.YAxis, AdjustedRad, Inventor.PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, Inventor.PartFeatureOperationEnum.kCutOperation)
							BackRevolveTrim.Name = FittingNameInput & " Upper trim revolution "
							BackRevolveTrim.SetAffectedBodies(TrimCol)
						
						
						If InputAngle = 0 Then 
						
							BackRevolveTrim.Suppressed = True
							RevolveTrim.Suppressed = False
						Else 	
							'Logger.Info("Inside the revolve trim suppression")
							RevolveTrim.Suppressed = False 'This should be turning on based on the input we have entered. 
							BackRevolveTrim.Suppressed = False
						End If