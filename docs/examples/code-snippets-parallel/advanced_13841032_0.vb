' Title: Selecting sketch profile on the assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/selecting-sketch-profile-on-the-assembly/td-p/13841032
' Category: advanced
' Scraped: 2025-10-09T08:59:00.259646

Dim oAsm As AssemblyDocument = ThisDoc.Document
Dim oDef As AssemblyComponentDefinition = oAsm.ComponentDefinition
Dim tg As TransientGeometry = ThisApplication.TransientGeometry

Dim skName As String = "Cut_Sketch"

' Parameters
Dim DSS As Double = Parameter("DSS")      ' diameter (mm)
Dim radius As Double = DSS / 2 / 10       ' convert mm → cm if model units = cm
Dim offsetDist As Double = ((DSS / 2) - 550) / 10

' 1️⃣ Find work plane P1
Dim basePlane As WorkPlane = Nothing
For Each wp As WorkPlane In oDef.WorkPlanes
    If wp.Name = "P1" Then
        basePlane = wp
        Exit For
    End If
Next
If basePlane Is Nothing Then
    MessageBox.Show("Work plane 'P1' not found.")
    Return
End If

' 2️⃣ Delete old sketch if exists
For Each s As PlanarSketch In oDef.Sketches
    If s.Name = skName Then
        Try : s.Delete() : Catch : End Try
        Exit For
    End If
Next

' 3️⃣ Create new sketch
Dim oSketch As PlanarSketch = oDef.Sketches.Add(basePlane)
oSketch.Name = skName

' 4️⃣ Get the origin of P1
Dim p1Origin As Inventor.Point
Try
    p1Origin = basePlane.Definition.PointOnPlane
Catch
    p1Origin = tg.CreatePoint(0, 0, 0)
End Try
Dim center2D As Point2d = oSketch.ModelToSketchSpace(p1Origin)

' 5️⃣ Draw circle (not construction)
Dim circle As SketchCircle = oSketch.SketchCircles.AddByCenterRadius(center2D, radius)
Dim dimRad As DimensionConstraint = oSketch.DimensionConstraints.AddRadius(circle, _
    tg.CreatePoint2d(center2D.X + radius * 0.5, center2D.Y + radius * 0.3))
dimRad.Parameter.Expression = "DSS/2"

' 6️⃣ Calculate line endpoints so they touch circle horizontally
Dim halfChord As Double = Math.Sqrt(radius ^ 2 - offsetDist ^ 2)
Dim start3D As Inventor.Point = tg.CreatePoint(-halfChord, 0, -offsetDist)
Dim end3D As Inventor.Point = tg.CreatePoint(halfChord, 0, -offsetDist)

Dim p1 As Point2d = oSketch.ModelToSketchSpace(start3D)
Dim p2 As Point2d = oSketch.ModelToSketchSpace(end3D)
Dim baseLine As SketchLine = oSketch.SketchLines.AddByTwoPoints(p1, p2)

' 7️⃣ Constrain line to be horizontal and coincident to circle
oSketch.GeometricConstraints.AddHorizontal(baseLine)
oSketch.GeometricConstraints.AddCoincident(baseLine.StartSketchPoint, circle)
oSketch.GeometricConstraints.AddCoincident(baseLine.EndSketchPoint, circle)

' 8️⃣ Add offset dimension (vertical distance)
Dim offsetDim As DimensionConstraint
offsetDim = oSketch.DimensionConstraints.AddTwoPointDistance( _
    circle.CenterSketchPoint, baseLine.StartSketchPoint, _
    DimensionOrientationEnum.kVerticalDim, _
    tg.CreatePoint2d(center2D.X + radius, center2D.Y - offsetDist * 0.8))
offsetDim.Parameter.Expression = "((DSS/2)-550)"

' 9️⃣ Ground circle center for stability
oSketch.GeometricConstraints.AddGround(circle.CenterSketchPoint)

'MessageBox.Show("Sketch created successfully: circle + line endpoints coincident to circle.")
'--- Corrected safe version ---
Dim oProfile As Profile
oProfile = oSketch.Profiles.AddForSolid(1)

Dim oProfiles As Profiles = oSketch.Profiles
If oProfiles.Count = 0 Then
    MessageBox.Show("No closed profiles found.")
    Return
End If



Dim selectedProfile As Profile = oProfiles.Item(1)




Dim oExtrudeDef As ExtrudeDefinition
oExtrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition( _
                selectedProfile, PartFeatureOperationEnum.kCutOperation)

oExtrudeDef.SetDistanceExtent(35, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)

Dim oCut As ExtrudeFeature
oCut = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)