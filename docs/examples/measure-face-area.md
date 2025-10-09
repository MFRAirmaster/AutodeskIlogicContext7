# iLogic Rule to Measure Area of Multiple Faces

This example demonstrates how to measure the surface area of named faces and track changes as sketches are modified.

## Inventor Version
Compatible with Inventor 2024 and later.

## Use Case
Measuring surface area of specific faces on a cylinder surface, useful for tracking changes in parametric designs where face sizes vary based on angles or other parameters.

## Description
This rule uses named entities to find specific faces, evaluates their surface area using SurfaceEvaluator, and can be combined with MODEL STATES to track area changes at different parameter values.

## Code
```vb
Dim iLogicAuto = iLogicVb.Automation

Dim namedEntities = iLogicAuto.GetNamedEntities(ThisDoc.Document)

Dim f As Face
f = namedEntities.FindEntity(COLD_OPENING), (HOT_OPENING)
Dim Eval As SurfaceEvaluator
Dim area As Double
Eval = f.Evaluator
area = Eval.Area
MsgBox("Face_Area : " & area)
```

## Key Concepts
- **Named Entities**: Use `iLogicVb.Automation.GetNamedEntities()` to access faces by name
- **SurfaceEvaluator**: `Face.Evaluator` provides geometric properties like Area
- **Model States**: Can be used to track area at different parameter values

## Usage Tips
1. Assign entity names to faces in the Inventor UI before running the rule
2. To measure multiple faces, call `FindEntity()` for each named face separately
3. Combine with POSITION parameters and MODEL STATES for parametric tracking
4. Area is returned in internal database units (cm²) - convert as needed

## Extended Example - Multiple Faces
```vb
Sub Main()
    Dim iLogicAuto = iLogicVb.Automation
    Dim namedEntities = iLogicAuto.GetNamedEntities(ThisDoc.Document)
    
    ' Measure first face
    Dim f1 As Face = namedEntities.FindEntity("COLD_OPENING")
    Dim area1 As Double = f1.Evaluator.Area
    
    ' Measure second face
    Dim f2 As Face = namedEntities.FindEntity("HOT_OPENING")
    Dim area2 As Double = f2.Evaluator.Area
    
    ' Calculate total area
    Dim totalArea As Double = area1 + area2
    
    ' Store in parameter
    Parameter.Param("Total_Surface_Area") = totalArea
    
    MsgBox("Total Area: " & totalArea & " cm²")
End Sub
