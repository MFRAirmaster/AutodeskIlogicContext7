# Part Rule with Reference Parameters to Top Assembly

This iLogic rule manipulates sketch dimensions in a part and is designed to be triggered from the top assembly level. It toggles the "driven" state of dimension constraints based on parameter values.

## Inventor Version
Compatible with Inventor 2024 and later (tested with 2025).

## Description
The rule accesses a specific sketch ("TANK SHELL") in the part, finds dimension constraints linked to parameters "TANK_EXTERNAL_DIA" and "TANK_INTERNAL_DIA", and toggles their driven state. It also updates assembly-level parameters.

## Code
```vb
Dim oPDoc As PartDocument = ThisDoc.Document
Dim oPDef As PartComponentDefinition = oPDoc.ComponentDefinition
Dim oSketch As PlanarSketch = oPDef.Sketches.Item("TANK SHELL")
Dim oDims As DimensionConstraints = oSketch.DimensionConstraints
Dim oHeightDim As DimensionConstraint
Dim oAngleDim As DimensionConstraint
For Each oDim As DimensionConstraint In oDims
    If oDim.Parameter.Name = "TANK_EXTERNAL_DIA" Then
        oHeightDim = oDim
    ElseIf oDim.Parameter.Name = "TANK_INTERNAL_DIA" Then
        oAngleDim = oDim
    End If
Next

TANK_EXTERNAL = Not TANK_INTERNAL

Try
    oAngleDim.Driven = Not oAngleDim.Driven
Catch
    oHeightDim.Driven = Not oHeightDim.Driven
    oAngleDim.Driven = Not oAngleDim.Driven
    Exit Sub
End Try
oHeightDim.Driven = Not oHeightDim.Driven
```

## Usage
Place this rule in the part file. To trigger from the assembly, use an assembly-level rule that calls this part rule via `iLogicVb.RunRule("PartName", "RuleName")`.
