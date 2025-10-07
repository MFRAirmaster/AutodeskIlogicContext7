# Parameter Manipulation Patterns

## Overview

This guide covers common patterns for manipulating parameters in iLogic, including calculations, validations, and dynamic updates.

## Basic Parameter Operations

### Reading and Writing Parameters

```vb
' Simple read and write
Dim currentLength As Double = Parameter("Length")
Parameter("Width") = 50.0

' Conditional update
If Parameter("Length") > 100 Then
    Parameter("Status") = "Oversized"
Else
    Parameter("Status") = "Normal"
End If
```

### Parameter Calculations

```vb
' Area calculation
Parameter("Area") = Parameter("Length") * Parameter("Width")

' Volume calculation for box
Parameter("Volume") = Parameter("Length") * Parameter("Width") * Parameter("Height")

' Perimeter calculation
Parameter("Perimeter") = 2 * (Parameter("Length") + Parameter("Width"))

' Diagonal calculation
Parameter("Diagonal") = Math.Sqrt(Parameter("Length")^2 + Parameter("Width")^2)
```

## Parameter Validation

### Range Validation

```vb
' Validate parameter is within range
Sub ValidateLength()
    Const MinLength As Double = 10.0
    Const MaxLength As Double = 1000.0
    
    Dim length As Double = Parameter("Length")
    
    If length < MinLength Then
        Parameter("Length") = MinLength
        MessageBox.Show("Length adjusted to minimum: " & MinLength, "Validation")
    ElseIf length > MaxLength Then
        Parameter("Length") = MaxLength
        MessageBox.Show("Length adjusted to maximum: " & MaxLength, "Validation")
    End If
End Sub
```

### Ratio Validation

```vb
' Maintain aspect ratio
Sub MaintainAspectRatio()
    Const TargetRatio As Double = 1.618  ' Golden ratio
    
    Dim length As Double = Parameter("Length")
    Dim width As Double = length / TargetRatio
    
    Parameter("Width") = width
End Sub
```

### Dependency Validation

```vb
' Ensure parameters maintain logical relationships
Sub ValidateDependencies()
    ' Thickness must be less than height
    If Parameter("Thickness") >= Parameter("Height") Then
        Parameter("Thickness") = Parameter("Height") * 0.1
        MessageBox.Show("Thickness adjusted to 10% of height", "Validation")
    End If
    
    ' Hole diameter must be less than width
    If Parameter("HoleDiameter") >= Parameter("Width") Then
        Parameter("HoleDiameter") = Parameter("Width") * 0.8
        MessageBox.Show("Hole diameter adjusted", "Validation")
    End If
End Sub
```

## Dynamic Parameter Updates

### Cascading Updates

```vb
' Update related parameters in cascade
Sub UpdateDimensions(baseSize As Double)
    ' Primary dimensions
    Parameter("Length") = baseSize
    Parameter("Width") = baseSize * 0.6
    Parameter("Height") = baseSize * 0.4
    
    ' Dependent dimensions
    Parameter("HoleDiameter") = baseSize * 0.15
    Parameter("HoleSpacing") = baseSize * 0.25
    Parameter("FilletRadius") = baseSize * 0.05
    
    ' Derived properties
    Parameter("Area") = Parameter("Length") * Parameter("Width")
    Parameter("Volume") = Parameter("Area") * Parameter("Height")
    
    iLogicVb.UpdateWhenDone = True
End Sub

' Call with base size
UpdateDimensions(100.0)
```

### Conditional Parameter Sets

```vb
' Apply different parameter sets based on configuration
Dim config As String = iProperties.Value("Custom", "Configuration")

Select Case config
    Case "Compact"
        Parameter("Length") = 50
        Parameter("Width") = 30
        Parameter("Height") = 20
        Parameter("FeatureEnabled") = 0
        
    Case "Standard"
        Parameter("Length") = 100
        Parameter("Width") = 60
        Parameter("Height") = 40
        Parameter("FeatureEnabled") = 1
        
    Case "Extended"
        Parameter("Length") = 200
        Parameter("Width") = 120
        Parameter("Height") = 80
        Parameter("FeatureEnabled") = 1
        Parameter("ExtraFeatures") = 1
End Select
```

## Parameter Arrays and Loops

### Working with Parameter Series

```vb
' Update series of similarly named parameters
For i As Integer = 1 To 10
    Parameter("Hole_" & i & "_Diameter") = 5.0 + (i * 0.5)
    Parameter("Hole_" & i & "_Depth") = 10.0 + (i * 1.0)
Next

' Loop through parameters to find maximum
Dim maxValue As Double = 0
For i As Integer = 1 To 5
    Dim currentValue As Double = Parameter("Section_" & i & "_Width")
    If currentValue > maxValue Then
        maxValue = currentValue
    End If
Next
Parameter("MaxWidth") = maxValue
```

### Batch Parameter Operations

```vb
' Apply operation to multiple parameters
Sub ScaleAllDimensions(scaleFactor As Double)
    Dim paramNames() As String = {"Length", "Width", "Height", "Diameter"}
    
    For Each paramName As String In paramNames
        If Parameter.Exists(paramName) Then
            Parameter(paramName) = Parameter(paramName) * scaleFactor
        End If
    Next
    
    iLogicVb.UpdateWhenDone = True
End Sub

' Scale up by 20%
ScaleAllDimensions(1.2)
```

## Mathematical Operations

### Trigonometric Calculations

```vb
' Calculate angle-based dimensions
Dim angleInDegrees As Double = Parameter("Angle")
Dim angleInRadians As Double = angleInDegrees * Math.PI / 180

' Calculate opposite side in right triangle
Dim adjacentSide As Double = Parameter("BaseLength")
Dim oppositeSide As Double = adjacentSide * Math.Tan(angleInRadians)
Parameter("Height") = oppositeSide

' Calculate hypotenuse
Dim hypotenuse As Double = adjacentSide / Math.Cos(angleInRadians)
Parameter("DiagonalLength") = hypotenuse
```

### Rounding and Precision

```vb
' Round to specific increments
Function RoundToIncrement(value As Double, increment As Double) As Double
    Return Math.Round(value / increment) * increment
End Function

' Round length to nearest 5mm
Parameter("Length") = RoundToIncrement(Parameter("RawLength"), 5.0)

' Round to 2 decimal places
Parameter("Price") = Math.Round(Parameter("RawPrice"), 2)

' Ceiling and floor operations
Parameter("MinSheets") = Math.Ceiling(Parameter("Area") / Parameter("SheetArea"))
Parameter("CompleteBoxes") = Math.Floor(Parameter("PartCount") / Parameter("BoxCapacity"))
```

## Conditional Parameter Logic

### Multi-Condition Updates

```vb
' Complex conditional logic
Dim material As String = iProperties.Value("Design Tracking", "Material")
Dim thickness As Double = Parameter("Thickness")
Dim length As Double = Parameter("Length")

' Determine load capacity based on multiple factors
Dim loadCapacity As Double

If material = "Steel" Then
    If thickness >= 5 Then
        loadCapacity = length * 100  ' High capacity
    Else
        loadCapacity = length * 50   ' Medium capacity
    End If
ElseIf material = "Aluminum" Then
    If thickness >= 8 Then
        loadCapacity = length * 60
    Else
        loadCapacity = length * 30
    End If
Else
    loadCapacity = length * 20  ' Default/unknown material
End If

Parameter("LoadCapacity") = loadCapacity
```

### State Machine Pattern

```vb
' Implement state-based parameter updates
Dim currentState As String = Parameter("State")

Select Case currentState
    Case "Initialized"
        Parameter("Length") = 100
        Parameter("State") = "Configured"
        
    Case "Configured"
        If Parameter("Length") > 0 Then
            Parameter("Area") = Parameter("Length") * Parameter("Width")
            Parameter("State") = "Calculated"
        End If
        
    Case "Calculated"
        If Parameter("Area") > 0 Then
            Parameter("MaterialNeeded") = Parameter("Area") * Parameter("Thickness") * Parameter("Density")
            Parameter("State") = "Complete"
        End If
        
    Case "Complete"
        ' All calculations done
        iLogicVb.UpdateWhenDone = True
End Select
```

## Parameter Interpolation

### Linear Interpolation

```vb
' Linear interpolation between two points
Function Lerp(start As Double, finish As Double, position As Double) As Double
    Return start + (finish - start) * position
End Function

' Use interpolation for smooth transitions
Dim transitionValue As Double = Parameter("TransitionPosition")  ' 0.0 to 1.0
Parameter("Length") = Lerp(50.0, 200.0, transitionValue)
Parameter("Width") = Lerp(30.0, 120.0, transitionValue)
```

### Table-Based Parameters

```vb
' Use lookup table for standard sizes
Function GetStandardSize(nominalSize As String) As Double
    Select Case nominalSize
        Case "XS"
            Return 50.0
        Case "S"
            Return 75.0
        Case "M"
            Return 100.0
        Case "L"
            Return 150.0
        Case "XL"
            Return 200.0
        Case Else
            Return 100.0  ' Default
    End Select
End Function

' Apply standard size
Dim sizeCode As String = iProperties.Value("Custom", "SizeCode")
Parameter("Length") = GetStandardSize(sizeCode)
```

## Performance Optimization

### Batch Updates

```vb
' Disable updates during batch operations
iLogicVb.UpdateWhenDone = False

' Perform all parameter updates
Parameter("Length") = 100
Parameter("Width") = 50
Parameter("Height") = 30
Parameter("Diameter") = 10
Parameter("Count") = 5

' Enable update once at the end
iLogicVb.UpdateWhenDone = True
```

### Conditional Calculation

```vb
' Only calculate when needed
Dim needsRecalculation As Boolean = Parameter("ForceRecalc") = 1

If needsRecalculation Then
    ' Expensive calculations here
    Dim complexValue As Double = 0
    For i As Integer = 1 To 1000
        complexValue += Math.Sqrt(i) * Parameter("Factor")
    Next
    Parameter("Result") = complexValue
    Parameter("ForceRecalc") = 0
End If
```

## Best Practices

### Error Handling

```vb
' Always use error handling for robust parameter operations
Try
    Dim value As Double = Parameter("Length")
    Parameter("Doubled") = value * 2
    
Catch ex As Exception
    MessageBox.Show("Error updating parameters: " & ex.Message, "Error")
    Logger.Error("Parameter update failed: " & ex.ToString())
End Try
```

### Parameter Existence Checks

```vb
' Always check if parameters exist before using them
If Parameter.Exists("OptionalParameter") Then
    Dim value As Double = Parameter("OptionalParameter")
    ' Use the value
Else
    ' Create the parameter or use default
    Parameter.UserParameters.AddByValue("OptionalParameter", 10.0, UnitsTypeEnum.kMillimeterLengthUnits)
End If
```

### Logging Parameter Changes

```vb
' Log important parameter changes for debugging
Sub UpdateAndLog(paramName As String, newValue As Double)
    Dim oldValue As Double = Parameter(paramName)
    Parameter(paramName) = newValue
    Logger.Info("Parameter '" & paramName & "' changed from " & oldValue & " to " & newValue)
End Sub
```

## Real-World Examples

### Example 1: Bolt Circle Pattern

```vb
' Calculate bolt hole positions on a circular pattern
Dim boltCircleDiameter As Double = Parameter("BoltCircleDiameter")
Dim numberOfHoles As Integer = CInt(Parameter("HoleCount"))
Dim angleIncrement As Double = 360.0 / numberOfHoles

For i As Integer = 0 To numberOfHoles - 1
    Dim angle As Double = i * angleIncrement * Math.PI / 180
    Dim xPos As Double = (boltCircleDiameter / 2) * Math.Cos(angle)
    Dim yPos As Double = (boltCircleDiameter / 2) * Math.Sin(angle)
    
    Parameter("Hole_" & (i + 1) & "_X") = xPos
    Parameter("Hole_" & (i + 1) & "_Y") = yPos
Next
```

### Example 2: Sheet Metal Bend Allowance

```vb
' Calculate bend allowance for sheet metal
Function CalculateBendAllowance(angle As Double, radius As Double, thickness As Double, kFactor As Double) As Double
    Dim angleRadians As Double = angle * Math.PI / 180
    Return angleRadians * (radius + (kFactor * thickness))
End Function

' Apply to parameters
Dim bendAngle As Double = Parameter("BendAngle")
Dim bendRadius As Double = Parameter("BendRadius")
Dim materialThickness As Double = Parameter("Thickness")
Dim kFactor As Double = 0.33  ' Typical for steel

Dim bendAllowance As Double = CalculateBendAllowance(bendAngle, bendRadius, materialThickness, kFactor)
Parameter("BendAllowance") = bendAllowance
Parameter("FlatLength") = Parameter("Leg1") + Parameter("Leg2") + bendAllowance
```

## Next Steps

- Learn about [Property Access Patterns](./02-property-access.md)
- Explore [Feature Control Patterns](./03-feature-control.md)
- See [API Reference](../api-reference/) for detailed object documentation
