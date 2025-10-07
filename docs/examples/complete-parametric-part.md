# Complete Parametric Part Example

## Overview

This comprehensive example demonstrates a real-world iLogic implementation for a configurable bracket part with multiple features, validations, and automated documentation.

## Part Description

A parametric mounting bracket with:
- Configurable dimensions
- Optional reinforcement ribs
- Variable hole patterns
- Material-dependent properties
- Automatic weight calculation
- Bill of materials integration

## Complete iLogic Rule

```vb
'================================================
' PARAMETRIC BRACKET CONFIGURATION RULE
' Purpose: Fully automated bracket design with validation
' Author: Engineering Team
' Version: 2.0
' Last Updated: 2025-01-07
'================================================

Try
    '=== CONFIGURATION SECTION ===
    
    ' Read configuration from iProperties
    Dim bracketSize As String = iProperties.Value("Custom", "BracketSize")
    Dim materialType As String = iProperties.Value("Design Tracking", "Material")
    Dim includeRibs As Boolean = (iProperties.Value("Custom", "IncludeRibs") = "Yes")
    Dim holePattern As String = iProperties.Value("Custom", "HolePattern")
    
    ' Log configuration
    Logger.Info("=== Bracket Configuration ===")
    Logger.Info("Size: " & bracketSize)
    Logger.Info("Material: " & materialType)
    Logger.Info("Ribs: " & includeRibs.ToString())
    Logger.Info("Hole Pattern: " & holePattern)
    
    
    '=== BASE DIMENSIONS ===
    
    ' Set base dimensions based on size category
    Select Case bracketSize
        Case "Small"
            Parameter("BaseLength") = 50.0
            Parameter("BaseWidth") = 40.0
            Parameter("BaseThickness") = 3.0
            Parameter("MountingHeight") = 30.0
            
        Case "Medium"
            Parameter("BaseLength") = 100.0
            Parameter("BaseWidth") = 80.0
            Parameter("BaseThickness") = 5.0
            Parameter("MountingHeight") = 60.0
            
        Case "Large"
            Parameter("BaseLength") = 150.0
            Parameter("BaseWidth") = 120.0
            Parameter("BaseThickness") = 8.0
            Parameter("MountingHeight") = 90.0
            
        Case "Custom"
            ' Use existing parameter values - no changes
            Logger.Info("Using custom dimensions")
            
        Case Else
            MessageBox.Show("Invalid bracket size. Using Medium.", "Warning")
            Parameter("BaseLength") = 100.0
            Parameter("BaseWidth") = 80.0
            Parameter("BaseThickness") = 5.0
            Parameter("MountingHeight") = 60.0
    End Select
    
    
    '=== DERIVED DIMENSIONS ===
    
    ' Calculate dependent dimensions
    Parameter("TotalHeight") = Parameter("MountingHeight") + Parameter("BaseThickness")
    Parameter("FilletRadius") = Parameter("BaseThickness") * 0.5
    Parameter("ChamferDistance") = Parameter("BaseThickness") * 0.3
    
    ' Calculate ribs dimensions (if enabled)
    If includeRibs Then
        Parameter("RibThickness") = Parameter("BaseThickness") * 0.6
        Parameter("RibHeight") = Parameter("MountingHeight") * 0.8
        Parameter("RibSpacing") = Parameter("BaseWidth") / 3
    End If
    
    
    '=== HOLE PATTERN CONFIGURATION ===
    
    Select Case holePattern
        Case "Single Center"
            Parameter("HoleCount") = 1
            Parameter("HoleDiameter") = 8.0
            Parameter("Hole_1_X") = 0.0
            Parameter("Hole_1_Y") = 0.0
            Feature.IsActive("LinearPattern") = False
            
        Case "Linear 3"
            Parameter("HoleCount") = 3
            Parameter("HoleDiameter") = 6.0
            Parameter("HoleSpacing") = Parameter("BaseWidth") / 4
            Feature.IsActive("LinearPattern") = True
            
        Case "Linear 5"
            Parameter("HoleCount") = 5
            Parameter("HoleDiameter") = 5.0
            Parameter("HoleSpacing") = Parameter("BaseWidth") / 6
            Feature.IsActive("LinearPattern") = True
            
        Case "Bolt Circle"
            Parameter("HoleCount") = 4
            Parameter("HoleDiameter") = 6.0
            Parameter("BoltCircleDiameter") = Parameter("BaseWidth") * 0.7
            
            ' Calculate positions for 4-hole bolt circle
            For i As Integer = 0 To 3
                Dim angle As Double = i * 90.0 * Math.PI / 180.0
                Dim radius As Double = Parameter("BoltCircleDiameter") / 2
                Parameter("Hole_" & (i + 1) & "_X") = radius * Math.Cos(angle)
                Parameter("Hole_" & (i + 1) & "_Y") = radius * Math.Sin(angle)
            Next
            
            Feature.IsActive("LinearPattern") = False
    End Select
    
    
    '=== FEATURE CONTROL ===
    
    ' Enable/disable reinforcement ribs
    Feature.IsActive("RibFeature1") = includeRibs
    Feature.IsActive("RibFeature2") = includeRibs
    
    ' Enable features based on size
    If bracketSize = "Large" Then
        Feature.IsActive("CornerGussets") = True
    Else
        Feature.IsActive("CornerGussets") = False
    End If
    
    
    '=== MATERIAL PROPERTIES ===
    
    ' Set material-dependent properties
    Dim density As Double
    Dim materialColor As String
    
    Select Case materialType
        Case "Steel"
            density = 7850.0  ' kg/m³
            materialColor = "Gray"
            Parameter("AllowableStress") = 250.0  ' MPa
            
        Case "Stainless Steel"
            density = 8000.0
            materialColor = "Silver"
            Parameter("AllowableStress") = 200.0
            
        Case "Aluminum"
            density = 2700.0
            materialColor = "Light Gray"
            Parameter("AllowableStress") = 110.0
            
        Case "Brass"
            density = 8500.0
            materialColor = "Gold"
            Parameter("AllowableStress") = 100.0
            
        Case Else
            density = 7850.0  ' Default to steel
            materialColor = "Gray"
            Parameter("AllowableStress") = 250.0
            MessageBox.Show("Unknown material. Using steel properties.", "Warning")
    End Select
    
    Parameter("MaterialDensity") = density
    
    
    '=== VALIDATION ===
    
    Dim validationErrors As New List(Of String)
    
    ' Validate dimensions
    If Parameter("BaseLength") < 20 Then
        validationErrors.Add("Base length must be at least 20mm")
    End If
    
    If Parameter("BaseWidth") < 20 Then
        validationErrors.Add("Base width must be at least 20mm")
    End If
    
    If Parameter("BaseThickness") < 2 Then
        validationErrors.Add("Base thickness must be at least 2mm")
    End If
    
    ' Validate hole spacing
    If holePattern.StartsWith("Linear") Then
        Dim minSpacing As Double = Parameter("HoleDiameter") * 2
        If Parameter("HoleSpacing") < minSpacing Then
            validationErrors.Add("Hole spacing too small. Minimum: " & minSpacing & "mm")
        End If
    End If
    
    ' Validate hole diameter
    Dim maxHoleDiameter As Double = Parameter("BaseWidth") * 0.3
    If Parameter("HoleDiameter") > maxHoleDiameter Then
        validationErrors.Add("Hole diameter too large. Maximum: " & maxHoleDiameter & "mm")
    End If
    
    ' Display validation errors
    If validationErrors.Count > 0 Then
        Dim errorMessage As String = "Validation Errors:" & vbCrLf & vbCrLf
        For Each errMsg As String In validationErrors
            errorMessage &= "• " & errMsg & vbCrLf
        Next
        MessageBox.Show(errorMessage, "Validation Failed")
        iLogicVb.UpdateWhenDone = False
        Exit Sub
    End If
    
    
    '=== CALCULATIONS ===
    
    ' Calculate volume (approximate)
    Dim baseVolume As Double = Parameter("BaseLength") * Parameter("BaseWidth") * Parameter("BaseThickness")
    Dim verticalVolume As Double = Parameter("BaseThickness") * Parameter("MountingHeight") * Parameter("BaseWidth")
    Dim totalVolume As Double = (baseVolume + verticalVolume) / 1000000.0  ' Convert mm³ to cm³
    
    ' Calculate mass
    Dim mass As Double = totalVolume * density / 1000.0  ' Convert to kg
    Parameter("CalculatedMass") = Math.Round(mass, 3)
    
    ' Calculate surface area (approximate)
    Dim surfaceArea As Double = 2 * (Parameter("BaseLength") * Parameter("BaseWidth") + _
                                      Parameter("BaseLength") * Parameter("TotalHeight") + _
                                      Parameter("BaseWidth") * Parameter("TotalHeight"))
    Parameter("SurfaceArea") = Math.Round(surfaceArea / 100.0, 2)  ' Convert to cm²
    
    ' Calculate material cost (example pricing)
    Dim costPerKg As Double = 5.0  ' Default cost
    Select Case materialType
        Case "Steel"
            costPerKg = 3.0
        Case "Stainless Steel"
            costPerKg = 12.0
        Case "Aluminum"
            costPerKg = 8.0
        Case "Brass"
            costPerKg = 15.0
    End Select
    
    Dim materialCost As Double = mass * costPerKg
    Parameter("MaterialCost") = Math.Round(materialCost, 2)
    
    
    '=== UPDATE iPROPERTIES ===
    
    ' Update part number based on configuration
    Dim partNumber As String = "BRK-" & bracketSize.Substring(0, 1) & "-" & materialType.Substring(0, 2).ToUpper()
    If includeRibs Then
        partNumber &= "-R"
    End If
    partNumber &= "-" & holePattern.Replace(" ", "")
    
    iProperties.Value("Project", "Part Number") = partNumber
    iProperties.Value("Project", "Description") = bracketSize & " " & materialType & " Mounting Bracket"
    iProperties.Value("Custom", "Mass") = mass.ToString("F3") & " kg"
    iProperties.Value("Custom", "MaterialCost") = "$" & materialCost.ToString("F2")
    iProperties.Value("Summary", "Comments") = "Auto-generated by iLogic on " & Now.ToString()
    
    
    '=== LOGGING ===
    
    Logger.Info("=== Calculation Results ===")
    Logger.Info("Volume: " & totalVolume.ToString("F2") & " cm³")
    Logger.Info("Mass: " & mass.ToString("F3") & " kg")
    Logger.Info("Surface Area: " & Parameter("SurfaceArea") & " cm²")
    Logger.Info("Material Cost: $" & materialCost.ToString("F2"))
    Logger.Info("Part Number: " & partNumber)
    Logger.Info("=== Configuration Complete ===")
    
    
    '=== FINAL UPDATE ===
    
    ' Update the model
    iLogicVb.UpdateWhenDone = True
    
    ' Success message
    MessageBox.Show("Bracket configured successfully!" & vbCrLf & vbCrLf & _
                   "Part Number: " & partNumber & vbCrLf & _
                   "Mass: " & mass.ToString("F3") & " kg" & vbCrLf & _
                   "Cost: $" & materialCost.ToString("F2"), _
                   "Configuration Complete")
    
Catch ex As Exception
    ' Error handling
    Dim errorMsg As String = "iLogic Error: " & ex.Message & vbCrLf & vbCrLf & _
                            "Stack Trace:" & vbCrLf & ex.StackTrace
    
    Logger.Error(errorMsg)
    MessageBox.Show(errorMsg, "iLogic Error")
    
    iLogicVb.UpdateWhenDone = False
End Try
```

## Required Parameters

Create these user parameters in your Inventor part:

### Primary Dimensions
- `BaseLength` (Double, mm) - Length of base plate
- `BaseWidth` (Double, mm) - Width of base plate
- `BaseThickness` (Double, mm) - Thickness of base plate
- `MountingHeight` (Double, mm) - Height of vertical mounting surface
- `TotalHeight` (Double, mm) - Total bracket height

### Hole Parameters
- `HoleCount` (Integer) - Number of mounting holes
- `HoleDiameter` (Double, mm) - Diameter of mounting holes
- `HoleSpacing` (Double, mm) - Spacing between holes
- `BoltCircleDiameter` (Double, mm) - Diameter of bolt circle (if applicable)
- `Hole_1_X` through `Hole_4_X` (Double, mm) - X positions of holes
- `Hole_1_Y` through `Hole_4_Y` (Double, mm) - Y positions of holes

### Optional Feature Parameters
- `RibThickness` (Double, mm) - Thickness of reinforcement ribs
- `RibHeight` (Double, mm) - Height of reinforcement ribs
- `RibSpacing` (Double, mm) - Spacing between ribs
- `FilletRadius` (Double, mm) - Radius of fillets
- `ChamferDistance` (Double, mm) - Distance of chamfers

### Calculated Parameters
- `CalculatedMass` (Double, kg) - Calculated part mass
- `SurfaceArea` (Double, cm²) - Surface area
- `MaterialCost` (Double, $) - Estimated material cost
- `MaterialDensity` (Double, kg/m³) - Material density
- `AllowableStress` (Double, MPa) - Material allowable stress

## Required iProperties

### Custom Properties
- `BracketSize` (Text) - Options: "Small", "Medium", "Large", "Custom"
- `IncludeRibs` (Text) - Options: "Yes", "No"
- `HolePattern` (Text) - Options: "Single Center", "Linear 3", "Linear 5", "Bolt Circle"
- `Mass` (Text) - Auto-populated
- `MaterialCost` (Text) - Auto-populated

### Design Tracking
- `Material` (Text) - Options: "Steel", "Stainless Steel", "Aluminum", "Brass"

## Required Features

Create these features in your Inventor part:
- `RibFeature1` - First reinforcement rib
- `RibFeature2` - Second reinforcement rib
- `LinearPattern` - Linear pattern of holes
- `CornerGussets` - Corner reinforcement gussets

## Usage Instructions

1. **Initial Setup**
   - Create a new Inventor part file
   - Add all required parameters (use snippet from Autodesk iLogic extension)
   - Add all required custom iProperties
   - Create the base features

2. **Configure the Part**
   - Set the `BracketSize` custom property
   - Set the `Material` in Design Tracking properties
   - Set `IncludeRibs` (Yes/No)
   - Set `HolePattern` option

3. **Run the Rule**
   - Execute the iLogic rule
   - The part will automatically configure all dimensions
   - Review the configuration summary message

4. **Validation**
   - Check for any validation error messages
   - Verify the part number in properties
   - Confirm mass and cost calculations

## Customization

### Adding New Sizes

To add a new size category, extend the Select Case statement:

```vb
Case "Extra Large"
    Parameter("BaseLength") = 200.0
    Parameter("BaseWidth") = 160.0
    Parameter("BaseThickness") = 10.0
    Parameter("MountingHeight") = 120.0
```

### Adding New Materials

Add material properties in the material selection section:

```vb
Case "Titanium"
    density = 4500.0
    materialColor = "Dark Gray"
    Parameter("AllowableStress") = 900.0
    costPerKg = 50.0
```

### Custom Hole Patterns

Create new hole pattern configurations:

```vb
Case "Cross Pattern"
    Parameter("HoleCount") = 4
    Parameter("HoleDiameter") = 6.0
    ' Set X,Y positions for each hole
    Parameter("Hole_1_X") = 10.0
    Parameter("Hole_1_Y") = 10.0
    ' ... and so on
```

## Best Practices Demonstrated

1. **Comprehensive Error Handling**: Try-Catch blocks protect against runtime errors
2. **Input Validation**: All inputs are validated before processing
3. **Logging**: Important events are logged for debugging
4. **Modular Design**: Functions and subroutines for reusable code
5. **Documentation**: Clear comments explain each section
6. **User Feedback**: MessageBox provides clear status updates
7. **iProperty Integration**: Automatic documentation updates
8. **Feature Control**: Dynamic enable/disable of optional features
9. **Calculated Properties**: Automatic mass and cost calculations
10. **Flexible Configuration**: Multiple configuration options

## Learning Points

- Complete parametric design workflow
- Integration of parameters, properties, and features
- Input validation and error handling
- Mathematical calculations for engineering
- Material property management
- Cost estimation
- Automatic documentation
- User-friendly configuration

## Next Steps

- Modify for your specific part requirements
- Add additional validation rules
- Integrate with Excel for material databases
- Add drawing automation
- Create assembly-level configuration rules

## Related Documentation

- [Basic Syntax](../core-concepts/01-basic-syntax.md)
- [Parameters and Properties](../core-concepts/02-parameters-properties.md)
- [Parameter Manipulation](../common-patterns/01-parameter-manipulation.md)
- [Best Practices](../best-practices/)
