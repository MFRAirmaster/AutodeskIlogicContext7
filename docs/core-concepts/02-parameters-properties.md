# Parameters and Properties in iLogic

## Overview

Parameters and properties are fundamental to iLogic programming. Parameters control model dimensions and behavior, while properties store metadata about documents.

## Model Parameters

### Accessing Parameters

```vb
' Read a parameter value
Dim length As Double = Parameter("Length")

' Set a parameter value
Parameter("Width") = 50.0

' Check if parameter exists
If Parameter.Exists("Height") Then
    Dim height As Double = Parameter("Height")
End If
```

### Parameter Types

iLogic works with different parameter types:

1. **Model Parameters**: Dimensions and user parameters
2. **Reference Parameters**: Linked from other documents
3. **Multi-value Parameters**: Parameters with lists of values

### Working with User Parameters

```vb
' Create a new user parameter
Parameter.UserParameters.AddByValue("TotalLength", 100.0, UnitsTypeEnum.kMillimeterLengthUnits)

' Delete a user parameter
If Parameter.Exists("TempParam") Then
    Parameter.Delete("TempParam")
End If

' Get parameter unit
Dim units As String = Parameter.Param("Length").Units

' Get parameter value with units
Dim lengthStr As String = Parameter("Length") & " mm"
```

### Parameter Units

```vb
' Common unit types
UnitsTypeEnum.kMillimeterLengthUnits    ' mm
UnitsTypeEnum.kCentimeterLengthUnits    ' cm
UnitsTypeEnum.kMeterLengthUnits         ' m
UnitsTypeEnum.kInchLengthUnits          ' inch
UnitsTypeEnum.kFootLengthUnits          ' foot
UnitsTypeEnum.kDegreeAngleUnits         ' degrees
UnitsTypeEnum.kRadianAngleUnits         ' radians

' Example: Create parameter with specific units
Parameter.UserParameters.AddByValue("Angle", 45.0, UnitsTypeEnum.kDegreeAngleUnits)
```

### Parameter Expressions

```vb
' Set parameter using expression
Parameter("TotalLength") = Parameter("Length1") + Parameter("Length2")

' Using equations
Parameter.Param("Area").Expression = "Length * Width"

' Get parameter expression
Dim expr As String = Parameter.Param("TotalCost").Expression
```

## iProperties

### Standard iProperties

```vb
' Read iProperty
Dim partNumber As String = iProperties.Value("Project", "Part Number")
Dim author As String = iProperties.Value("Summary", "Author")
Dim material As String = iProperties.Value("Design Tracking", "Material")

' Write iProperty
iProperties.Value("Project", "Part Number") = "PN-12345"
iProperties.Value("Summary", "Title") = "Bracket Assembly"
iProperties.Value("Summary", "Comments") = "Updated by iLogic"
```

### iProperty Categories

| Category | Common Properties |
|----------|------------------|
| `Summary` | Title, Subject, Author, Comments, Keywords |
| `Project` | Part Number, Project, Designer, Cost Center |
| `Status` | Design Status, Checked By, Date Checked |
| `Custom` | User-defined properties |
| `Design Tracking` | Material, Part Number, Stock Number |
| `Physical` | Mass, Area, Volume, Density |

### Custom iProperties

```vb
' Create or update custom iProperty
iProperties.Value("Custom", "CustomerName") = "Acme Corp"
iProperties.Value("Custom", "OrderNumber") = "ORD-2024-001"
iProperties.Value("Custom", "Revision") = "A"

' Read custom iProperty
Dim customer As String = iProperties.Value("Custom", "CustomerName")

' Check if custom property exists
Try
    Dim value As String = iProperties.Value("Custom", "NonExistent")
Catch
    MessageBox.Show("Property does not exist")
End Try
```

### Physical Properties

```vb
' Read physical properties (read-only)
Dim mass As Double = iProperties.Mass  ' in kg
Dim area As Double = iProperties.Area  ' in cm²
Dim volume As Double = iProperties.Volume  ' in cm³

' Display physical properties
MessageBox.Show("Mass: " & mass & " kg" & vbCrLf & _
                "Volume: " & volume & " cm³", _
                "Physical Properties")
```

## Multi-Value Parameters

Multi-value parameters allow you to define lists of valid values for a parameter.

### Creating Multi-Value Lists

```vb
' Define multi-value list for material
MultiValue.List("Material") = "Steel;Aluminum;Brass;Plastic"

' Or using array
Dim materials() As String = {"Steel", "Aluminum", "Brass", "Plastic"}
MultiValue.List("Material") = String.Join(";", materials)
```

### Using Multi-Value Parameters

```vb
' Get current value
Dim currentMaterial As String = MultiValue.Value("Material")

' Set value (must be from the list)
MultiValue.Value("Material") = "Aluminum"

' Get all possible values
Dim valuesList As String = MultiValue.List("Material")
Dim valuesArray() As String = valuesList.Split(";"c)

' Check if value is in list
If MultiValue.List("Material").Contains("Titanium") Then
    MessageBox.Show("Titanium is available")
End If
```

### Dynamic Multi-Value Lists

```vb
' Create dynamic list based on another parameter
Dim materialType As String = Parameter("MaterialType")

Select Case materialType
    Case "Metal"
        MultiValue.List("Material") = "Steel;Aluminum;Brass;Copper"
    Case "Plastic"
        MultiValue.List("Material") = "ABS;Nylon;Polycarbonate;PVC"
    Case "Wood"
        MultiValue.List("Material") = "Oak;Pine;Maple;Birch"
End Select
```

## Parameter Collections

### Iterating Through Parameters

```vb
' Loop through all model parameters
For Each param In ThisDoc.Document.ComponentDefinition.Parameters.ModelParameters
    MessageBox.Show("Parameter: " & param.Name & " = " & param.Value)
Next

' Loop through user parameters
For Each param In ThisDoc.Document.ComponentDefinition.Parameters.UserParameters
    Logger.Info("User Parameter: " & param.Name & " = " & param.Value)
Next
```

### Finding Parameters

```vb
' Find parameters by name pattern
For Each param In ThisDoc.Document.ComponentDefinition.Parameters.ModelParameters
    If param.Name.StartsWith("Hole") Then
        Logger.Info("Found hole parameter: " & param.Name)
    End If
Next
```

## Component Parameters

### Assembly Component Parameters

```vb
' Access component parameter in assembly
Dim compLength As Double = Component.InventorComponent("SubAssembly:1").Parameter("Length")

' Set component parameter
Component.InventorComponent("Part1:1").Parameter("Width") = 25.0

' Check if component parameter exists
If Component.InventorComponent("Part1:1").ParameterExists("Height") Then
    Dim height As Double = Component.InventorComponent("Part1:1").Parameter("Height")
End If
```

### Component iProperties

```vb
' Access component iProperty
Dim partNum As String = Component.InventorComponent("Part1:1").GetPropertyValue("Part Number")

' Set component iProperty
Component.InventorComponent("Part1:1").SetPropertyValue("Part Number", "PN-001")
```

## Parameter Naming Conventions

### Best Practices

```vb
' Good naming conventions
Parameter("OverallLength")      ' Clear, descriptive
Parameter("HoleDiameter_1")     ' Numbered series
Parameter("Shaft_Length")       ' Underscore separation
Parameter("isEnabled")          ' Boolean prefix

' Avoid
Parameter("x")                  ' Too vague
Parameter("Length123")          ' Unclear meaning
Parameter("My Parameter")       ' Spaces (works but not recommended)
```

## Practical Examples

### Example 1: Conditional Parameter Update

```vb
' Update dimensions based on size category
Dim sizeCategory As String = iProperties.Value("Custom", "Size")

Select Case sizeCategory
    Case "Small"
        Parameter("Length") = 50
        Parameter("Width") = 30
        Parameter("Height") = 20
    Case "Medium"
        Parameter("Length") = 100
        Parameter("Width") = 60
        Parameter("Height") = 40
    Case "Large"
        Parameter("Length") = 200
        Parameter("Width") = 120
        Parameter("Height") = 80
End Select

iLogicVb.UpdateWhenDone = True
```

### Example 2: Parameter Validation

```vb
' Validate parameter values
Function ValidateParameters() As Boolean
    Dim isValid As Boolean = True
    
    ' Check minimum length
    If Parameter("Length") < 10 Then
        MessageBox.Show("Length must be at least 10mm", "Validation Error")
        isValid = False
    End If
    
    ' Check aspect ratio
    Dim aspectRatio As Double = Parameter("Length") / Parameter("Width")
    If aspectRatio > 10 Then
        MessageBox.Show("Length/Width ratio cannot exceed 10:1", "Validation Error")
        isValid = False
    End If
    
    Return isValid
End Function

' Use validation
If ValidateParameters() Then
    iLogicVb.UpdateWhenDone = True
Else
    iLogicVb.UpdateWhenDone = False
End If
```

### Example 3: Bulk Parameter Update

```vb
' Update multiple related parameters
Sub UpdateHoleDimensions(diameter As Double)
    Parameter("HoleDiameter") = diameter
    Parameter("HoleDepth") = diameter * 2
    Parameter("CounterboreDiameter") = diameter * 1.5
    Parameter("CounterboreDepth") = diameter * 0.5
    
    ' Log changes
    Logger.Info("Updated hole dimensions for diameter: " & diameter)
End Sub

' Call the subroutine
UpdateHoleDimensions(8.0)
iLogicVb.UpdateWhenDone = True
```

### Example 4: Property-Driven Configuration

```vb
' Configure part based on iProperties
Dim config As String = iProperties.Value("Custom", "Configuration")
Dim revision As String = iProperties.Value("Project", "Revision")

Select Case config
    Case "Standard"
        Parameter("FeatureCount") = 4
        Feature.IsActive("ExtraHoles") = False
    Case "Heavy Duty"
        Parameter("FeatureCount") = 6
        Feature.IsActive("ExtraHoles") = True
        Parameter("Thickness") = 10
    Case "Light Weight"
        Parameter("FeatureCount") = 2
        Feature.IsActive("ExtraHoles") = False
        Parameter("Thickness") = 3
End Select

' Update part number with revision
Dim basePartNum As String = iProperties.Value("Project", "Part Number")
iProperties.Value("Project", "Stock Number") = basePartNum & "-Rev" & revision
```

## Common Pitfalls

1. **Case Sensitivity**: Parameter names are case-sensitive
```vb
' These reference different parameters
Parameter("Length")  
Parameter("length")  ' Different from above!
```

2. **Units**: Always be aware of unit conversions
```vb
' Length in mm
Dim lengthMM As Double = Parameter("Length")

' Convert to inches
Dim lengthInch As Double = lengthMM / 25.4
```

3. **Circular References**: Avoid circular parameter dependencies
```vb
' BAD - Circular reference
Parameter("A") = Parameter("B")
Parameter("B") = Parameter("A")  ' Will cause error
```

## Next Steps

- Learn about [Rules and External Rules](./03-rules-external-rules.md)
- Explore [Event Triggers](./04-event-triggers.md)
- See [Common Patterns](../common-patterns/) for more examples
