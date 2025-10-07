# Extended iLogic Examples Collection

This document contains additional iLogic code examples and variations based on forum discussions and common programming patterns.

## Table of Contents

### 🔄 [Parameter Operations](#parameter-operations)
- [Bulk Parameter Updates](#bulk-parameter-updates)
- [Parameter Validation](#parameter-validation)
- [Dynamic Parameter Creation](#dynamic-parameter-creation)

### 📐 [Geometry Automation](#geometry-automation)
- [Advanced Sketch Operations](#advanced-sketch-operations)
- [Feature Creation Patterns](#feature-creation-patterns)
- [Multi-Body Operations](#multi-body-operations)

### 🔗 [Assembly Automation](#assembly-automation)
- [Component Management](#component-management)
- [Constraint Automation](#constraint-automation)
- [Bill of Materials](#bill-of-materials)

### 🎨 [User Interface](#user-interface)
- [Custom Forms](#custom-forms)
- [Progress Indicators](#progress-indicators)
- [Error Handling](#error-handling)

---

## Parameter Operations

### Bulk Parameter Updates

**Update Multiple Parameters:**
```vb
' Bulk parameter update with validation
Dim paramsToUpdate As New Dictionary(Of String, Double) From {
    {"Length", 100.0},
    {"Width", 50.0},
    {"Height", 25.0},
    {"Thickness", 2.0}
}

For Each param In paramsToUpdate
    Try
        Parameter(param.Key) = param.Value
        Logger.Info($"Updated {param.Key} to {param.Value}")
    Catch ex As Exception
        MessageBox.Show($"Error updating {param.Key}: {ex.Message}", "Parameter Update Error")
    End Try
Next

iLogicVb.UpdateWhenDone = True
```

**Conditional Parameter Updates:**
```vb
' Update parameters based on conditions
Dim materialType As String = Parameter("Material_Type")

Select Case materialType.ToUpper()
    Case "STEEL"
        Parameter("Density") = 7850 ' kg/m³
        Parameter("Yield_Strength") = 250 ' MPa
    Case "ALUMINUM"
        Parameter("Density") = 2700
        Parameter("Yield_Strength") = 95
    Case "TITANIUM"
        Parameter("Density") = 4500
        Parameter("Yield_Strength") = 880
    Case Else
        MessageBox.Show("Unknown material type: " & materialType, "Material Error")
        Return
End Select

' Calculate weight
Parameter("Volume") = Parameter("Length") * Parameter("Width") * Parameter("Height") / 1000000000 ' Convert mm³ to m³
Parameter("Weight") = Parameter("Volume") * Parameter("Density")
```

### Parameter Validation

**Range Validation:**
```vb
' Parameter range validation
Dim length As Double = Parameter("Length")
Dim minLength As Double = 10
Dim maxLength As Double = 1000

If length < minLength Or length > maxLength Then
    MessageBox.Show($"Length must be between {minLength} and {maxLength} mm", "Validation Error")
    Parameter("Length") = Math.Max(minLength, Math.Min(maxLength, length))
    Return
End If

' Check relationships
If Parameter("Width") > Parameter("Length") Then
    MessageBox.Show("Width cannot be greater than length", "Validation Error")
    Parameter("Width") = Parameter("Length") * 0.8
End If
```

**Dependency Validation:**
```vb
' Validate parameter dependencies
Dim holeDiameter As Double = Parameter("Hole_Diameter")
Dim materialThickness As Double = Parameter("Thickness")

If holeDiameter >= materialThickness Then
    MessageBox.Show("Hole diameter cannot be larger than material thickness", "Design Error")
    Parameter("Hole_Diameter") = materialThickness * 0.8
End If

' Check clearance
Dim shaftDiameter As Double = Parameter("Shaft_Diameter")
Dim clearance As Double = (holeDiameter - shaftDiameter) / 2

If clearance < 0.1 Then
    MessageBox.Show("Insufficient clearance between shaft and hole", "Clearance Warning")
End If
```

### Dynamic Parameter Creation

**Runtime Parameter Creation:**
```vb
' Create parameters dynamically
Dim paramNames() As String = {"Custom1", "Custom2", "Custom3"}
Dim paramValues() As Double = {10.5, 20.3, 15.7}

For i As Integer = 0 To paramNames.Length - 1
    Try
        ' Check if parameter exists
        Dim testValue As Double = Parameter(paramNames(i))
    Catch ex As Exception
        ' Parameter doesn't exist, create it
        Parameter(paramNames(i)) = paramValues(i)
        Logger.Info($"Created parameter {paramNames(i)} = {paramValues(i)}")
    End Try
Next
```

---

## Geometry Automation

### Advanced Sketch Operations

**Parametric Circle with Construction Lines:**
```vb
Dim oSketch As PlanarSketch = oDef.Sketches.Add(oWorkPlane)
oSketch.Name = "Parametric_Circle"

' Get sketch center
Dim centerPoint As Point2d = oSketch.ModelToSketchSpace(oCenterPoint)

' Create construction lines
Dim hLine As SketchLine = oSketch.SketchLines.AddByTwoPoints(
    tg.CreatePoint2d(centerPoint.X - 50, centerPoint.Y),
    tg.CreatePoint2d(centerPoint.X + 50, centerPoint.Y))
hLine.Construction = True

Dim vLine As SketchLine = oSketch.SketchLines.AddByTwoPoints(
    tg.CreatePoint2d(centerPoint.X, centerPoint.Y - 50),
    tg.CreatePoint2d(centerPoint.X, centerPoint.Y + 50))
vLine.Construction = True

' Create circle
Dim circle As SketchCircle = oSketch.SketchCircles.AddByCenterRadius(centerPoint, Parameter("Radius"))

' Add dimensions
Dim radiusDim As DimensionConstraint = oSketch.DimensionConstraints.AddRadius(circle,
    tg.CreatePoint2d(centerPoint.X + Parameter("Radius"), centerPoint.Y))
radiusDim.Parameter.Expression = "Radius"

' Add constraints
oSketch.GeometricConstraints.AddCoincident(circle.CenterSketchPoint, hLine)
oSketch.GeometricConstraints.AddCoincident(circle.CenterSketchPoint, vLine)
```

**Polygon Creation:**
```vb
' Create regular polygon
Dim sides As Integer = Parameter("Sides")
Dim radius As Double = Parameter("Radius")

Dim oSketch As PlanarSketch = oDef.Sketches.Add(oWorkPlane)
Dim center As Point2d = oSketch.ModelToSketchSpace(oCenterPoint)

Dim points As New List(Of Point2d)
For i As Integer = 0 To sides - 1
    Dim angle As Double = (2 * Math.PI * i) / sides
    Dim x As Double = center.X + radius * Math.Cos(angle)
    Dim y As Double = center.Y + radius * Math.Sin(angle)
    points.Add(tg.CreatePoint2d(x, y))
Next

' Create polygon lines
Dim polygonLines As New List(Of SketchLine)
For i As Integer = 0 To points.Count - 1
    Dim nextIndex As Integer = (i + 1) Mod points.Count
    Dim line As SketchLine = oSketch.SketchLines.AddByTwoPoints(points(i), points(nextIndex))
    polygonLines.Add(line)
Next

' Add constraints
For Each line As SketchLine In polygonLines
    oSketch.GeometricConstraints.AddEqualLength(line, polygonLines(0))
Next
```

### Feature Creation Patterns

**Extrude with Draft:**
```vb
' Create extrude with draft angle
Dim oProfile As Profile = oSketch.Profiles.AddForSolid()

Dim extrudeDef As ExtrudeDefinition
extrudeDef = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, PartFeatureOperationEnum.kJoinOperation)

' Set distance and draft
extrudeDef.SetDistanceExtent(Parameter("Height"), PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
extrudeDef.SetDraftAngle(Parameter("Draft_Angle"), PartFeatureExtentDirectionEnum.kPositiveExtentDirection)

' Create feature
Dim extrudeFeature As ExtrudeFeature = oDef.Features.ExtrudeFeatures.Add(extrudeDef)
extrudeFeature.Name = "Drafted_Extrusion"
```

**Hole Pattern Creation:**
```vb
' Create hole pattern
Dim holeCenter As Point2d = oSketch.ModelToSketchSpace(oHoleLocation)
Dim holeCircle As SketchCircle = oSketch.SketchCircles.AddByCenterRadius(holeCenter, Parameter("Hole_Radius")/2)

Dim holeProfile As Profile = oSketch.Profiles.AddForSolid()
Dim holeExtrude As ExtrudeDefinition = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(holeProfile, PartFeatureOperationEnum.kCutOperation)
holeExtrude.SetDistanceExtent(Parameter("Thickness") + 10, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)

Dim holeFeature As ExtrudeFeature = oDef.Features.ExtrudeFeatures.Add(holeExtrude)

' Create rectangular pattern
Dim patternDef As RectangularPatternDefinition
patternDef = oDef.Features.RectangularPatternFeatures.CreateDefinition()

patternDef.ParentFeatures.Add(holeFeature)
patternDef.XDirectionEntity = oXAxis
patternDef.YDirectionEntity = oYAxis
patternDef.XCount = Parameter("X_Count")
patternDef.YCount = Parameter("Y_Count")
patternDef.XSpacing = Parameter("X_Spacing")
patternDef.YSpacing = Parameter("Y_Spacing")

Dim patternFeature As RectangularPatternFeature = oDef.Features.RectangularPatternFeatures.Add(patternDef)
```

### Multi-Body Operations

**Body Separation:**
```vb
' Split part into multiple bodies
Dim splitTool As WorkPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oBasePlane, Parameter("Split_Height"))

Dim splitDef As SplitDefinition = oDef.Features.SplitFeatures.CreateDefinition(oTargetBody)
splitDef.SplitTool = splitTool

Dim splitFeature As SplitFeature = oDef.Features.SplitFeatures.Add(splitDef)

' Name the resulting bodies
Dim bodies As SurfaceBodies = oDef.SurfaceBodies
For i As Integer = 1 To bodies.Count
    bodies.Item(i).Name = $"Body_{i}"
Next
```

---

## Assembly Automation

### Component Management

**Dynamic Component Replacement:**
```vb
' Replace component based on parameter
Dim currentComponent As ComponentOccurrence = FindComponentByName(oAsmCompDef, "Widget")

If currentComponent IsNot Nothing Then
    Dim newComponentPath As String = ""

    Select Case Parameter("Widget_Type").ToString().ToUpper()
        Case "TYPE_A"
            newComponentPath = "C:\Library\Widget_A.ipt"
        Case "TYPE_B"
            newComponentPath = "C:\Library\Widget_B.ipt"
        Case "TYPE_C"
            newComponentPath = "C:\Library\Widget_C.ipt"
    End Select

    If newComponentPath <> "" Then
        currentComponent.Replace(newComponentPath, True)
    End If
End If
```

**Component Visibility Control:**
```vb
' Control component visibility
Dim components As ComponentOccurrences = oAsmCompDef.Occurrences

For Each comp As ComponentOccurrence In components
    Select Case comp.Name.ToLower()
        Case "optional_part_1"
            comp.Visible = Parameter("Show_Optional_1")
        Case "optional_part_2"
            comp.Visible = Parameter("Show_Optional_2")
        Case "variant_a"
            comp.Visible = (Parameter("Variant") = "A")
        Case "variant_b"
            comp.Visible = (Parameter("Variant") = "B")
    End Select
Next
```

### Constraint Automation

**Adaptive Constraints:**
```vb
' Create adaptive assembly constraints
Dim comp1 As ComponentOccurrence = FindComponentByName(oAsmCompDef, "Part1")
Dim comp2 As ComponentOccurrence = FindComponentByName(oAsmCompDef, "Part2")

If comp1 IsNot Nothing And comp2 IsNot Nothing Then
    ' Get faces for mating
    Dim face1 As Face = FindFaceByName(comp1, "Mate_Face_1")
    Dim face2 As Face = FindFaceByName(comp2, "Mate_Face_2")

    If face1 IsNot Nothing And face2 IsNot Nothing Then
        ' Create mate constraint
        Dim mateDef As MateConstraintDefinition
        mateDef = oAsmCompDef.Constraints.CreateMateConstraintDefinition()

        mateDef.EntityOne = face1
        mateDef.EntityTwo = face2
        mateDef.Offset.Value = Parameter("Mate_Offset")

        Dim mateConstraint As MateConstraint = oAsmCompDef.Constraints.AddMateConstraint(mateDef)
        mateConstraint.Name = "Adaptive_Mate"
    End If
End If
```

### Bill of Materials

**BOM Customization:**
```vb
' Customize BOM data
Dim oBOM As BOM = oAsmCompDef.BOM
oBOM.PartsOnlyViewEnabled = True

Dim oBOMView As BOMView = oBOM.BOMViews.Item("Parts Only")

For Each oBOMRow As BOMRow In oBOMView.BOMRows
    Dim oCompDef As ComponentDefinition = oBOMRow.ComponentDefinitions.Item(1)

    ' Set custom part number
    If oCompDef.Type = ObjectTypeEnum.kPartComponentDefinitionObject Then
        Dim partDef As PartComponentDefinition = oCompDef
        oBOMRow.ItemNumber = GetCustomPartNumber(partDef)
    End If

    ' Set quantity override
    If Parameter("Override_Qty") Then
        oBOMRow.TotalQuantity = Parameter("Custom_Qty")
    End If
Next
```

---

## User Interface

### Custom Forms

**Dynamic Form Creation:**
```vb
' Create custom input form
Dim formTitle As String = "Design Parameters"
Dim form As New InputForm(formTitle)

' Add controls dynamically
form.AddTextBox("Length", "Length (mm)", Parameter("Length"))
form.AddTextBox("Width", "Width (mm)", Parameter("Width"))
form.AddComboBox("Material", "Material", New String() {"Steel", "Aluminum", "Plastic"}, Parameter("Material"))

' Add checkboxes
form.AddCheckBox("Include_Holes", "Include mounting holes", Parameter("Include_Holes"))
form.AddCheckBox("Surface_Finish", "Apply surface finish", Parameter("Surface_Finish"))

' Show form
form.Show()

' Process results
If form.OK Then
    Parameter("Length") = form.GetDouble("Length")
    Parameter("Width") = form.GetDouble("Width")
    Parameter("Material") = form.GetString("Material")
    Parameter("Include_Holes") = form.GetBoolean("Include_Holes")
    Parameter("Surface_Finish") = form.GetBoolean("Surface_Finish")

    iLogicVb.UpdateWhenDone = True
End If
```

### Progress Indicators

**Progress Bar for Long Operations:**
```vb
' Show progress for batch operations
Dim progressForm As New ProgressForm("Processing Components...")
progressForm.Show()

Dim totalComponents As Integer = oAsmCompDef.Occurrences.Count
Dim currentComponent As Integer = 0

For Each comp As ComponentOccurrence In oAsmCompDef.Occurrences
    currentComponent += 1

    ' Update progress
    Dim progress As Double = (currentComponent / totalComponents) * 100
    progressForm.UpdateProgress(progress, $"Processing {comp.Name}...")

    ' Perform operation
    ProcessComponent(comp)

    ' Allow UI updates
    System.Windows.Forms.Application.DoEvents()
Next

progressForm.Close()
```

### Error Handling

**Comprehensive Error Handling:**
```vb
Try
    ' Main operation
    PerformComplexOperation()

Catch ex As ArgumentException
    MessageBox.Show($"Invalid parameter: {ex.Message}", "Parameter Error")

Catch ex As IOException
    MessageBox.Show($"File operation failed: {ex.Message}", "File Error")

Catch ex As InvalidOperationException
    MessageBox.Show($"Operation not allowed: {ex.Message}", "Operation Error")

Catch ex As Exception
    ' Generic error handler
    MessageBox.Show($"Unexpected error: {ex.Message}{vbCrLf}{ex.StackTrace}", "Error")

    ' Log error for debugging
    Logger.Error($"Error in rule: {ex.ToString()}")

Finally
    ' Cleanup operations
    CleanupResources()

End Try
```

**Retry Logic:**
```vb
Dim maxRetries As Integer = 3
Dim retryCount As Integer = 0
Dim success As Boolean = False

While Not success And retryCount < maxRetries
    Try
        retryCount += 1

        ' Attempt operation
        PerformUnreliableOperation()
        success = True

    Catch ex As Exception
        Logger.Warning($"Attempt {retryCount} failed: {ex.Message}")

        If retryCount < maxRetries Then
            System.Threading.Thread.Sleep(1000) ' Wait 1 second
        Else
            MessageBox.Show($"Operation failed after {maxRetries} attempts: {ex.Message}", "Error")
        End If
    End Try
End While
```

---

## Additional Resources

- [Forum Examples](forum-scraped-examples.md)
- [Advanced Examples](forum-advanced-examples.md)
- [Comprehensive Guide](comprehensive-ilogic-guide.md)

---

*These examples extend the basic patterns with more advanced techniques and error handling.*
*Last Updated: 2025-01-07*
