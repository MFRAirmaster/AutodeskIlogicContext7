# Comprehensive iLogic Programming Guide

This guide provides a complete reference for iLogic programming in Autodesk Inventor, compiled from community forum examples and best practices.

## Table of Contents

### 🔰 [Basic Concepts](#basic-concepts)
- [Parameters and Properties](#parameters-and-properties)
- [Basic Syntax](#basic-syntax)
- [Simple Automation](#simple-automation)

### 🛠️ [Common Patterns](#common-patterns)
- [Geometry Operations](#geometry-operations)
- [Assembly Automation](#assembly-automation)
- [File Operations](#file-operations)

### 🎯 [Advanced Techniques](#advanced-techniques)
- [API Integration](#api-integration)
- [Multi-Body Operations](#multi-body-operations)
- [Event-Driven Programming](#event-driven-programming)

### 📚 [Code Examples](#code-examples)
- [Forum Sourced Examples](#forum-sourced-examples)
- [Advanced API Examples](#advanced-api-examples)

---

## Basic Concepts

### Parameters and Properties

**Setting Parameters:**
```vb
' Set a parameter value
Parameter("Length") = 100.0
Parameter("Width") = 50.0

' Use parameters in expressions
Parameter("Area") = Parameter("Length") * Parameter("Width")
```

**Working with iProperties:**
```vb
' Set custom iProperties
iProperties.Value("Custom", "Part_Number") = "PN-001"
iProperties.Value("Project", "Description") = "Widget Component"

' Get iProperties
Dim partNum As String = iProperties.Value("Custom", "Part_Number")
```

### Basic Syntax

**Variable Declaration:**
```vb
' Explicit typing
Dim length As Double = 100.0
Dim partName As String = "Component1"
Dim isActive As Boolean = True

' Type inference
Dim width = 50.0  ' Double
Dim count = 10     ' Integer
```

**Control Structures:**
```vb
' If-Then-Else
If Parameter("Length") > 100 Then
    Parameter("Size") = "Large"
ElseIf Parameter("Length") > 50 Then
    Parameter("Size") = "Medium"
Else
    Parameter("Size") = "Small"
End If

' For loops
For i As Integer = 1 To 10
    Parameter("Instance_" & i) = i * 10
Next

' While loops
Dim counter As Integer = 0
While counter < 10
    counter = counter + 1
    MessageBox.Show("Count: " & counter)
End While
```

### Simple Automation

**Update Model:**
```vb
' Force model update
iLogicVb.UpdateWhenDone = True

' Suppress update temporarily
InventorVb.DocumentUpdate()

' Manual update
ThisDoc.Document.Update()
```

---

## Common Patterns

### Geometry Operations

**Face Area Measurement:**
```vb
Dim iLogicAuto = iLogicVb.Automation
Dim namedEntities = iLogicAuto.GetNamedEntities(ThisDoc.Document)

Dim f As Face
f = namedEntities.FindEntity("FACE_NAME")
Dim Eval As SurfaceEvaluator = f.Evaluator
Dim area As Double = Eval.Area

MessageBox.Show("Face Area: " & area & " mm²")
```

**Sketch Creation:**
```vb
Dim oSketch As PlanarSketch = oDef.Sketches.Add(oWorkPlane)
oSketch.Name = "My_Sketch"

' Add geometry
Dim center As Point2d = oSketch.ModelToSketchSpace(oCenterPoint)
Dim circle As SketchCircle = oSketch.SketchCircles.AddByCenterRadius(center, radius)

' Add constraints
oSketch.GeometricConstraints.AddGround(circle.CenterSketchPoint)
```

### Assembly Automation

**Component Selection:**
```vb
' Select components in assembly
oDoc.SelectSet.Select(Component.InventorComponent(MakePath("75430:2", "75424:1")))

' Clear selection
oDoc.SelectSet.Clear()
```

**Work Plane Creation:**
```vb
Dim oWP As WorkPlane = oDef.WorkPlanes.AddByPlaneAndOffset(oBasePlane, offset)
oWP.Name = "Custom_Plane"
```

### File Operations

**Save Document:**
```vb
' Save current document
ThisDoc.Save()

' Save as new file
ThisDoc.Document.SaveAs("C:\Path\NewFile.ipt", True)
```

**Open Document:**
```vb
Dim oDoc As Document = ThisApplication.Documents.Open("C:\Path\File.ipt")
```

---

## Advanced Techniques

### API Integration

**ControlDefinition Execution:**
```vb
' Execute Inventor commands programmatically
Dim oControlDef As Inventor.ControlDefinition
oControlDef = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(False)  ' Asynchronous execution

' Alternative for triggers
oControlDef.Execute2(True)   ' Synchronous execution
```

**Assembly Constraints:**
```csharp
// C# API example for assembly constraints
AssemblyDocument asmDoc = (AssemblyDocument)InventorApp.ActiveDocument;
AssemblyComponentDefinition oAsmCompDef = asmDoc.ComponentDefinition;

// Create geometry proxies
object wpx1, wpx2;
compOcc.CreateGeometryProxy(compDef.WorkPlanes["XY Plane"], out wpx1);

// Add constraint
AssemblyConstraints AssCons = oAsmCompDef.Constraints;
AssCons.AddFlushConstraint((WorkPlaneProxy)wpx1, (WorkPlaneProxy)wpx2, 0);
```

### Multi-Body Operations

**RevolveFeature with Multi-Body:**
```vb
' Specify which solid body to affect
Dim revolveDef As RevolveDefinition
revolveDef = oCompDef.Features.RevolveFeatures.CreateRevolveDefinition(oProfile, oAxis, PartFeatureOperationEnum.kCutOperation)

' Set angle extent
revolveDef.SetAngleExtent(angle, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)

' Specify affected bodies
revolveDef.AffectedBodies.Add(targetBody)

' Create feature
Dim revolveFeat As RevolveFeature = oCompDef.Features.RevolveFeatures.Add(revolveDef)
```

### Event-Driven Programming

**After Open Document Trigger:**
```vb
' Code that runs when document opens
MessageBox.Show("Document opened: " & ThisDoc.FileName)

' Automatic operations
If Parameter("Auto_Update") = True Then
    iLogicVb.UpdateWhenDone = True
End If
```

---

## Code Examples

### Forum Sourced Examples

#### Measuring Face Area
```vb
Dim iLogicAuto = iLogicVb.Automation
Dim namedEntities = iLogicAuto.GetNamedEntities(ThisDoc.Document)
Dim f As Face = namedEntities.FindEntity("COLD_OPENING")
Dim Eval As SurfaceEvaluator = f.Evaluator
Dim area As Double = Eval.Area
MsgBox("Face_Area : " & area)
```

#### SelectSet with MakePath
```vb
' Select nested component
oDoc.SelectSet.Select(Component.InventorComponent(MakePath("75430:2", "75424:1")))

' Select first-level component
oDoc.SelectSet.Select(Component.InventorComponent("75430:2"))
```

#### ControlDefinition Execute
```vb
MessageBox.Show("Updating Revision Table")
Dim oControlDef As Inventor.ControlDefinition = ThisApplication.CommandManager.ControlDefinitions.Item("VaultRevisionBlockTableNode")
oControlDef.Execute2(True)
MessageBox.Show("Updated Revision Table")
```

#### 3D Sketch Between Work Points
```vb
Dim doc As PartDocument = ThisDoc.Document
Dim def As PartComponentDefinition = doc.ComponentDefinition

Dim wpA1 = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kWorkPointFilter, "Select work point A1")
Dim wpB1 = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kWorkPointFilter, "Select work point B1")

Dim sketch3D As Sketch3D = def.Sketches3D.Add()
sketch3D.SketchLines3D.AddByTwoPoints(wpA1, wpB1)
sketch3D.Name = ("Strut 01")
```

### Advanced API Examples

#### Assembly Sketch Profile Selection
```vb
Dim oAsm As AssemblyDocument = ThisDoc.Document
Dim oDef As AssemblyComponentDefinition = oAsm.ComponentDefinition
Dim tg As TransientGeometry = ThisApplication.TransientGeometry

' Find work plane
Dim basePlane As WorkPlane = Nothing
For Each wp As WorkPlane In oDef.WorkPlanes
    If wp.Name = "P1" Then
        basePlane = wp
        Exit For
    End If
Next

' Create sketch and geometry
Dim oSketch As PlanarSketch = oDef.Sketches.Add(basePlane)
Dim center2D As Point2d = oSketch.ModelToSketchSpace(p1Origin)
Dim circle As SketchCircle = oSketch.SketchCircles.AddByCenterRadius(center2D, radius)

' Create extrude cut
Dim oProfiles As Profiles = oSketch.Profiles
Dim selectedProfile As Profile = oProfiles.Item(1)
Dim oExtrudeDef As ExtrudeDefinition = oDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(selectedProfile, PartFeatureOperationEnum.kCutOperation)
oExtrudeDef.SetDistanceExtent(35, PartFeatureExtentDirectionEnum.kPositiveExtentDirection)
Dim oCut As ExtrudeFeature = oDef.Features.ExtrudeFeatures.Add(oExtrudeDef)
```

#### RevolveFeature for Cutting Solids
```vb
Dim revolveFeatRemove As RevolveFeature = DefinizioneParte.Features.RevolveFeatures.AddByAngle( _
    oProfile01, oWorkAx, oElbowAngle, _
    PartFeatureExtentDirectionEnum.kNegativeExtentDirection, _
    PartFeatureOperationEnum.kCutOperation)
```

---

## Best Practices

### Code Organization
- Use meaningful variable names
- Add comments for complex logic
- Group related operations
- Use consistent indentation

### Error Handling
```vb
Try
    ' Risky operation
    Dim result = SomeOperation()
Catch ex As Exception
    MessageBox.Show("Error: " & ex.Message)
End Try
```

### Performance Optimization
- Minimize document updates
- Use batch operations when possible
- Avoid unnecessary loops
- Cache frequently used objects

### Debugging Techniques
- Use MessageBox for variable inspection
- Add logging with Debug.Print
- Test with simple cases first
- Use breakpoints in external rules

---

## Resources

- [Forum Examples](forum-scraped-examples.md)
- [Advanced Examples](forum-advanced-examples.md)
- [Core Concepts](../core-concepts/01-basic-syntax.md)
- [API Reference](../api-reference/)

---

*This guide is continuously updated with community contributions and new examples.*
*Last Updated: 2025-01-07*
