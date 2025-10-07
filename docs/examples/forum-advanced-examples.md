# Advanced iLogic Forum Examples

This document contains advanced code examples scraped from the Autodesk Inventor Programming Forum, focusing on complex scenarios and API usage.

## Table of Contents
- [Assembly Sketch Profile Selection](#assembly-sketch-profile-selection)
- [Assembly Constraints with Origin Planes](#assembly-constraints-with-origin-planes)
- [RevolveFeature for Cutting Solids](#revolvefeature-for-cutting-solids)

---

## Assembly Sketch Profile Selection

**Source:** [Forum Post](https://forums.autodesk.com/t5/inventor-programming-forum/selecting-sketch-profile-on-the-assembly/td-p/13841032)

**Problem:** Creating an extrude cut on an assembly with automatic profile selection. The challenge is selecting a specific profile region (e.g., right part of a circle) rather than all closed profiles.

**Complete Code Example:**

```vb
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

' 🔟 Select profile and create cut extrude
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
```

**Key Concepts:**

1. **Assembly Sketching**: Creating sketches in assembly context using work planes
2. **TransientGeometry**: Using transient geometry for calculations
3. **Profile Selection**: Working with sketch profiles and selecting specific regions
4. **Geometric Constraints**: 
   - `AddHorizontal()` - Horizontal constraint
   - `AddCoincident()` - Coincidence between points and curves
   - `AddGround()` - Fixed/grounded geometry
5. **Dimension Constraints**: 
   - `AddRadius()` - Radius dimension
   - `AddTwoPointDistance()` - Distance between two points with orientation
6. **Parameter Expressions**: Linking dimensions to model parameters
7. **Model-to-Sketch Space**: Converting between 3D model coordinates and 2D sketch coordinates

**Use Cases:**
- Parametric assembly cuts
- Automated feature creation based on design parameters
- Complex profile selection for manufacturing features

**Common Issues:**
- Profile selection may capture all closed regions; use index selection carefully
- Ensure sketch is fully constrained for robust parametric behavior
- Unit conversion between mm and cm based on document settings

---

## Assembly Constraints with Origin Planes

**Source:** [Forum Post](https://forums.autodesk.com/t5/inventor-programming-forum/how-to-create-a-constraint-with-a-part-and-the-assembly-s-origin/td-p/13840352)

**Problem:** Creating assembly constraints between a component and the assembly's origin planes using the Inventor API (C#).

**Code Example (C#):**

```csharp
// Get active assembly document
AssemblyDocument asmDoc = (AssemblyDocument)Globals.InventorApp.ActiveDocument;
AssemblyComponentDefinition oAsmCompDef = asmDoc.ComponentDefinition;

// Adds a component (in this case a part) to the assembly.
ComponentOccurrence compOcc = AddComponent(PartPath);
// Get the part's definition
PartComponentDefinition compDef = (PartComponentDefinition)compOcc.Definition;

// Create proxies for the workplanes
object wpx1, wpx2;
compOcc.CreateGeometryProxy(compDef.WorkPlanes["XY Plane"], out wpx1);

// Creating the proxy for the assembly's origin XY Plane causes a COMException.
// "System.Runtime.InteropServices.COMException: 'Exception has been thrown by the target of an invocation.'"
compOcc.CreateGeometryProxy(oAsmCompDef.WorkPlanes["XY Plane"], out wpx2);

// Add the constraint to the assembly. This is never reached as an exception is thrown when creating the proxy.
AssemblyConstraints AssCons = oAsmCompDef.Constraints;
AssCons.AddFlushConstraint((WorkPlaneProxy)wpx1, (WorkPlaneProxy)wpx2, 0);
```

**Key Concepts:**

1. **Assembly API (C#)**: Working with assemblies in C# add-ins
2. **Component Occurrences**: Managing component instances in assemblies
3. **Geometry Proxies**: Creating proxies for constraint references
4. **Assembly Constraints**: Adding flush constraints between planes

**Issue Identified:**
The code throws a COMException when attempting to create a geometry proxy for the assembly's origin plane through a component occurrence.

**Solution Approach:**
Assembly origin planes don't need proxies since they exist at the assembly level. Instead:

```csharp
// For assembly origin planes, use them directly:
WorkPlane asmXYPlane = oAsmCompDef.WorkPlanes["XY Plane"];

// Create proxy only for component geometry
object wpx1;
compOcc.CreateGeometryProxy(compDef.WorkPlanes["XY Plane"], out wpx1);

// Add constraint using component proxy and direct assembly plane reference
AssemblyConstraints AssCons = oAsmCompDef.Constraints;
AssCons.AddFlushConstraint((WorkPlaneProxy)wpx1, asmXYPlane, 0);
```

**Use Cases:**
- Automated component placement
- Assembly automation
- Fixture and jig creation
- Standard component positioning

---

## RevolveFeature for Cutting Solids

**Source:** [Forum Post](https://forums.autodesk.com/t5/inventor-programming-forum/cut-a-solid-by-revolvefeature/td-p/13837083)

**Problem:** Selecting specific solid bodies to cut when using RevolveFeature.

**Code Example:**

```vb
Dim revolveFeatRemove As RevolveFeature = DefinizioneParte.Features.RevolveFeatures.AddByAngle( _
    oProfile01, _
    oWorkAx, _
    oElbowAngle, _
    PartFeatureExtentDirectionEnum.kNegativeExtentDirection, _
    PartFeatureOperationEnum.kCutOperation)
```

**Key Concepts:**

1. **RevolveFeature**: Creating revolved features (cuts, joins, intersects)
2. **Profile Selection**: Using sketch profiles for revolve operations
3. **Work Axis**: Defining the axis of revolution
4. **Extent Direction**: Controlling revolve direction
5. **Feature Operation**: Cut, join, or intersect operations

**Enhanced Solution for Multi-Body Selection:**

```vb
' For multi-body parts, specify which solid to affect
Dim oPartDoc As PartDocument = ThisApplication.ActiveDocument
Dim oCompDef As PartComponentDefinition = oPartDoc.ComponentDefinition

' Get the solid body you want to cut
Dim targetBody As SurfaceBody = oCompDef.SurfaceBodies.Item(1) ' or find by name

' Create revolve definition
Dim revolveD as RevolveDefinition
revolveDef = oCompDef.Features.RevolveFeatures.CreateRevolveDefinition( _
    oProfile01, _
    oWorkAx, _
    PartFeatureOperationEnum.kCutOperation)

' Set the angle
revolveDef.SetAngleExtent(oElbowAngle, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)

' Specify which solid body to affect (for multi-body parts)
revolveDef.AffectedBodies.Add(targetBody)

' Create the revolve feature
Dim revolveFeatRemove As RevolveFeature = oCompDef.Features.RevolveFeatures.Add(revolveDef)
```

**Use Cases:**
- Creating rotational cuts (holes, grooves, chamfers)
- Pipe and tube modifications
- Lathe-type operations
- Symmetric feature removal

**Advanced Tips:**
1. Use `RevolveDefinition` for better control over affected bodies
2. In multi-body parts, always specify `AffectedBodies`
3. Validate profile is closed and properly constrained
4. Ensure work axis doesn't intersect the profile (for full revolves)

---

## Additional Resources

For more examples:
- [Basic Forum Examples](forum-scraped-examples.md)
- [Core Concepts](../core-concepts/01-basic-syntax.md)
- [API Reference](../api-reference/)

---

*Last Updated: 2025-01-07*
*These are advanced examples requiring good understanding of Inventor API*
