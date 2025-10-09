' Title: Sheet metal cut feature affected bodies
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sheet-metal-cut-feature-affected-bodies/td-p/13843101
' Category: advanced
' Scraped: 2025-10-09T08:54:18.006574

' =====================================
' 🔧 USER SETTINGS
' =====================================
Dim FilePath As String = "Template.xlsx"
Dim SheetName As String = "Sheet1"

Dim oPart As PartDocument = ThisApplication.ActiveDocument
Dim oDef As SheetMetalComponentDefinition = oPart.ComponentDefinition
Dim oTG As TransientGeometry = ThisApplication.TransientGeometry
oPart.Activate()

' --- Read how many rows to process ---
Dim RowCount As Integer = CInt(GoExcel.CellValue(FilePath, SheetName, "B2"))

' --- Pre-cleanup: remove extra WorkPoints, sketches, planes ---
For Each wp As WorkPoint In oDef.WorkPoints
    If wp.Name.Contains("NP_") Then wp.Delete()
Next
For Each sk As PlanarSketch In oDef.Sketches
    If sk.Name.Contains("Sketch_NP_") Then sk.Delete()
Next
For Each wp As WorkPlane In oDef.WorkPlanes
    If wp.Name.Contains("Plane_NP_") Then wp.Delete()
Next

' =====================================
' --- LOOP ---
' =====================================
For i As Integer = 5 To 5 + RowCount - 1
    ' --- Read Excel values ---
    Dim PointName As String = "NP_" & GoExcel.CellValue(FilePath, SheetName, "A" & i)
    Dim Y As Double = -CDbl(GoExcel.CellValue(FilePath, SheetName, "B" & i)) / 10
    Dim Distance As Double = CDbl(GoExcel.CellValue(FilePath, SheetName, "C" & i)) / 10
    Dim AngleDeg As Double = CDbl(GoExcel.CellValue(FilePath, SheetName, "D" & i)) - 90
    Dim Diameter As Double = CDbl(GoExcel.CellValue(FilePath, SheetName, "E" & i)) / 10

    Dim AngleRad As Double = AngleDeg * Math.PI / 180
    Dim X As Double = Distance * Math.Cos(AngleRad)
    Dim Z As Double = Distance * Math.Sin(AngleRad)

    ' --- Delete existing points ---
    For Each wp As WorkPoint In oDef.WorkPoints
        If wp.Name = PointName Or wp.Name = PointName & "_direction" Then wp.Delete()
    Next

    ' --- Create points ---
    Dim oMainPoint As WorkPoint = oDef.WorkPoints.AddFixed(oTG.CreatePoint(X, Y, Z))
    oMainPoint.Name = PointName
    oMainPoint.Visible = True

    Dim oDirPoint As WorkPoint = oDef.WorkPoints.AddFixed(oTG.CreatePoint(0, Y, 0))
    oDirPoint.Name = PointName & "_direction"
    oDirPoint.Visible = True

    ' --- Create axis ---
    Dim AxisName As String = "Axis_" & PointName
    For Each ax As WorkAxis In oDef.WorkAxes
        If ax.Name = AxisName Then ax.Delete()
    Next
    Dim oAxis As WorkAxis = oDef.WorkAxes.AddByTwoPoints(oDirPoint, oMainPoint)
    oAxis.Name = AxisName
    oAxis.Visible = False

    ' --- Create plane ---
    Dim PlaneName As String = "Plane_" & PointName
    For Each wp As WorkPlane In oDef.WorkPlanes
        If wp.Name = PlaneName Then wp.Delete()
    Next
    Dim oPlane As WorkPlane = oDef.WorkPlanes.AddByNormalToCurve(oAxis, oMainPoint)
    oPlane.Name = PlaneName
    oPlane.Visible = True

    ' --- Create sketch ---
    Dim SketchName As String = "Sketch_" & PointName
    For Each sk As PlanarSketch In oDef.Sketches
        If sk.Name = SketchName Then sk.Delete()
    Next
    Dim oSketch As PlanarSketch = oDef.Sketches.Add(oPlane)
    oSketch.Name = SketchName

    ' --- Circle ---
    Dim oSketchPoint As SketchPoint = oSketch.AddByProjectingEntity(oMainPoint)
    Dim oCircle As SketchCircle = oSketch.SketchCircles.AddByCenterRadius(oSketchPoint, Diameter / 2)
    Try
        oSketch.GeometricConstraints.AddCoincident(oCircle.CenterSketchPoint, oSketchPoint)
    Catch
    End Try

    ' --- Add diameter dimension (optional)
    Dim textPt As Point2d = oTG.CreatePoint2d(Diameter + 5, 0)
    oSketch.DimensionConstraints.AddDiameter(oCircle, textPt)

	' --- CUT FEATURE ---
	Dim oSMFeatures As SheetMetalFeatures = oDef.Features
	
	' --- Create cut definition ---
	Dim oCutDef As CutDefinition = oSMFeatures.CutFeatures.CreateCutDefinition(oSketch.Profiles.AddForSolid())
	
	' --- Set cut distance using the Distance variable from Excel ---
	oCutDef.SetDistanceExtent(Distance, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
	oCutDef.CutNormalToFlat = True
	
	' --- Add the cut ---
	Dim oCut As CutFeature = oSMFeatures.CutFeatures.Add(oCutDef)
	oCut.Name = "Cut_" & PointName
Next