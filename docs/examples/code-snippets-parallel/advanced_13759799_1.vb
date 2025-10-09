' Title: How get Center point in iLogic sketch?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-get-center-point-in-ilogic-sketch/td-p/13759799#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:58:29.668195

sub createSkeletonBox()
	'===============Parameters===============
	Dim Vyska = createUserParameter("Vyska", 600, "mm").Value
	Dim Sirka = createUserParameter("Sirka", 450, "mm").Value
	Dim Hloubka = createUserParameter("Hloubka", 400, "mm").Value

	Dim dX = Sirka/2
	Dim dY = Hloubka

	'===============Definitions===============
	Dim oCompDef as PartComponentDefinition
	oCompDef = ThisDoc.Document.ComponentDefinition

	Dim tg As TransientGeometry
	tg = ThisApplication.TransientGeometry

	'===============Work===============

	'===Create skeletonSketch===
	Dim oSketch As PlanarSketch
	oSketch = oCompDef.Sketches.Add(oCompDef.WorkPlanes(3))
	oSketch.name = "skeletonSketch"
	
	Dim lines As Inventor.SketchLines = oSketch.SketchLines
	Dim points As Inventor.SketchPoints = oSketch.SketchPoints
	Dim constr As Inventor.GeometricConstraints = oSketch.GeometricConstraints

	Dim CenterPoint = points.Add(tg.CreatePoint2d(0, 0), False) '??????????????
	'===Points===
	Dim point_array(3) As Inventor.SketchPoint
		point_array(0) = points.Add(tg.CreatePoint2d(0, 0), False)
		point_array(1) = points.Add(tg.CreatePoint2d(4, 0), False)
		point_array(2) = points.Add(tg.CreatePoint2d(4, 4), False)
		point_array(3) = points.Add(tg.CreatePoint2d(0, 4), False)

	'===Draw skeletonSketch===
		Dim line0 = lines.AddByTwoPoints(point_array(0), point_array(1))
		Dim line1 = lines.AddByTwoPoints(point_array(1), point_array(2))
		Dim line2 = lines.AddByTwoPoints(point_array(2), point_array(3))
		Dim line3 = lines.AddByTwoPoints(point_array(3), point_array(0))

	'===Constraints===
	constr.AddHorizontal(line0)
	constr.AddHorizontal(line2)
	constr.AddVertical(line1)
	constr.AddVertical(line3)

	'===Dimensions===
	oSketch.DimensionConstraints.AddTwoPointDistance(point_array(0), point_array(1), DimensionOrientationEnum.kHorizontalDim, tg.CreatePoint2d(0, -2)).Parameter.Expression = "Sirka"
	oSketch.DimensionConstraints.AddTwoPointDistance(point_array(1), point_array(2), DimensionOrientationEnum.kVerticalDim, tg.CreatePoint2d(-2, 0)).Parameter.Expression = "Hloubka"
	'oSketch.DimensionConstraints.AddTwoPointDistance(point_array(0), CenterPoint, DimensionOrientationEnum.kHorizontalDim, tg.CreatePoint2d(0, -2)).Parameter.Expression = dX
	'oSketch.DimensionConstraints.AddTwoPointDistance(point_array(0), CenterPoint, DimensionOrientationEnum.kVerticalDim, tg.CreatePoint2d(0, -2)).Parameter.Expression = dY
	
	'===Create a profile===
	Dim oProfile As Profile
    oProfile = oSketch.Profiles.AddForSolid

	'===Create a base extrusion===
	'Dim oExtrude As ExtrudeFeature
	'oExtrude = oCompDef.Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, "Vyska", kPositiveExtentDirection, kNewBodyOperation)
	'oExtrude.SurfaceBodies(1).Name = "@Skeleton"
	'oExtrude.name = "skeletonExtrude"

end sub