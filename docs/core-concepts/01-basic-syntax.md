# Basic iLogic Syntax

## Overview

iLogic uses Visual Basic .NET (VB.NET) syntax for writing rules. This guide covers the fundamental syntax elements you'll use in iLogic programming.

## Comments

```vb
' Single-line comment

' Multiple single-line comments
' can be used for longer explanations
```

## Variables and Data Types

### Declaring Variables

```vb
' Implicit typing (not recommended)
Dim value = 100

' Explicit typing (recommended)
Dim length As Double = 100.0
Dim partName As String = "Bracket"
Dim isValid As Boolean = True
Dim count As Integer = 5
```

### Common Data Types

| Type | Description | Example |
|------|-------------|---------|
| `Integer` | Whole numbers | `Dim count As Integer = 10` |
| `Double` | Decimal numbers | `Dim length As Double = 25.4` |
| `String` | Text | `Dim name As String = "Part1"` |
| `Boolean` | True/False | `Dim flag As Boolean = True` |
| `Date` | Date and time | `Dim today As Date = Now` |
| `Object` | Any object | `Dim doc As Object = ThisDoc.Document` |

## Operators

### Arithmetic Operators

```vb
Dim result As Double

result = 10 + 5   ' Addition: 15
result = 10 - 5   ' Subtraction: 5
result = 10 * 5   ' Multiplication: 50
result = 10 / 5   ' Division: 2
result = 10 ^ 2   ' Exponentiation: 100
result = 10 Mod 3 ' Modulus: 1
```

### Comparison Operators

```vb
Dim a As Integer = 10
Dim b As Integer = 5

If a = b Then   ' Equal to
If a <> b Then  ' Not equal to
If a > b Then   ' Greater than
If a < b Then   ' Less than
If a >= b Then  ' Greater than or equal
If a <= b Then  ' Less than or equal
```

### Logical Operators

```vb
Dim x As Boolean = True
Dim y As Boolean = False

If x And y Then      ' Logical AND
If x Or y Then       ' Logical OR
If Not x Then        ' Logical NOT
If x AndAlso y Then  ' Short-circuit AND
If x OrElse y Then   ' Short-circuit OR
```

## String Operations

```vb
' Concatenation
Dim fullName As String = "Part" & "_" & "001"  ' "Part_001"

' String interpolation (using & operator)
Dim length As Double = 100
Dim message As String = "Length is " & length & " mm"

' Common string methods
Dim text As String = "  Hello World  "
Dim upper As String = text.ToUpper()        ' "  HELLO WORLD  "
Dim lower As String = text.ToLower()        ' "  hello world  "
Dim trimmed As String = text.Trim()         ' "Hello World"
Dim length As Integer = text.Length         ' 15
Dim contains As Boolean = text.Contains("Hello")  ' True
```

## Conditional Statements

### If-Then-Else

```vb
' Simple If
If length > 100 Then
    MessageBox.Show("Length exceeds maximum")
End If

' If-Else
If length > 100 Then
    MessageBox.Show("Too long")
Else
    MessageBox.Show("Length OK")
End If

' If-ElseIf-Else
If length > 100 Then
    MessageBox.Show("Too long")
ElseIf length < 50 Then
    MessageBox.Show("Too short")
Else
    MessageBox.Show("Length OK")
End If

' Inline If (Ternary operator)
Dim status As String = If(length > 100, "Invalid", "Valid")
```

### Select Case

```vb
Dim materialType As String = iProperties.Value("Custom", "Material")

Select Case materialType
    Case "Steel"
        density = 7850
    Case "Aluminum"
        density = 2700
    Case "Plastic"
        density = 1200
    Case Else
        density = 0
        MessageBox.Show("Unknown material")
End Select
```

## Loops

### For Loop

```vb
' Basic For loop
For i As Integer = 1 To 10
    MessageBox.Show("Iteration: " & i)
Next

' For loop with Step
For i As Integer = 0 To 100 Step 10
    Parameter("Value_" & i) = i * 2
Next

' For loop counting down
For i As Integer = 10 To 1 Step -1
    MessageBox.Show(i)
Next
```

### For Each Loop

```vb
' Loop through collection
Dim paramList As New List(Of String)
paramList.Add("Length")
paramList.Add("Width")
paramList.Add("Height")

For Each paramName As String In paramList
    MessageBox.Show("Parameter: " & paramName)
Next
```

### While Loop

```vb
Dim counter As Integer = 0

While counter < 10
    counter = counter + 1
    Parameter("Counter") = counter
End While
```

### Do-While Loop

```vb
Dim value As Double = 0

Do While value < 100
    value = value + 10
    Parameter("Value") = value
Loop

' Alternative: Do-Until
Do Until value >= 100
    value = value + 10
Loop
```

## Functions and Subroutines

### Functions (Return a Value)

```vb
' Define a function
Function CalculateArea(length As Double, width As Double) As Double
    Return length * width
End Function

' Use the function
Dim area As Double = CalculateArea(100, 50)
Parameter("Area") = area
```

### Subroutines (No Return Value)

```vb
' Define a subroutine
Sub UpdateDimensions(newLength As Double, newWidth As Double)
    Parameter("Length") = newLength
    Parameter("Width") = newWidth
End Sub

' Call the subroutine
UpdateDimensions(100, 50)
```

## Arrays and Collections

### Arrays

```vb
' Declare array
Dim numbers(4) As Integer  ' Array of 5 elements (0-4)

' Initialize array
Dim values() As Double = {10.0, 20.0, 30.0, 40.0, 50.0}

' Access elements
values(0) = 15.0
Dim firstValue As Double = values(0)

' Loop through array
For Each value As Double In values
    MessageBox.Show(value.ToString())
Next
```

### Lists (Generic Collections)

```vb
' Create a list
Dim partNumbers As New List(Of String)

' Add items
partNumbers.Add("PN-001")
partNumbers.Add("PN-002")
partNumbers.Add("PN-003")

' Access items
Dim first As String = partNumbers(0)
Dim count As Integer = partNumbers.Count

' Remove items
partNumbers.Remove("PN-002")
partNumbers.RemoveAt(0)

' Check if contains
If partNumbers.Contains("PN-001") Then
    MessageBox.Show("Part number found")
End If
```

## Error Handling

```vb
Try
    ' Code that might cause an error
    Dim value As Double = Parameter("NonExistentParam")
    
Catch ex As Exception
    ' Handle the error
    MessageBox.Show("Error: " & ex.Message, "Error")
    
Finally
    ' Code that always runs (optional)
    ' Clean up resources here
End Try
```

## Best Practices

1. **Always use explicit typing**: `Dim value As Double = 100.0`
2. **Use meaningful variable names**: `partLength` instead of `x`
3. **Comment your code**: Explain complex logic
4. **Handle errors appropriately**: Use Try-Catch blocks
5. **Use constants for fixed values**: `Const MaxLength As Double = 1000.0`

## Example: Complete Rule

```vb
' Rule: Calculate and update part dimensions
' Author: Engineering Team
' Date: 2025-01-07

Try
    ' Get input parameters
    Dim baseLength As Double = Parameter("BaseLength")
    Dim scaleFactor As Double = Parameter("ScaleFactor")
    
    ' Validate inputs
    If scaleFactor <= 0 Then
        MessageBox.Show("Scale factor must be positive", "Validation Error")
        Exit Sub
    End If
    
    ' Calculate new length
    Dim newLength As Double = baseLength * scaleFactor
    
    ' Check constraints
    Const MaxLength As Double = 1000.0
    If newLength > MaxLength Then
        MessageBox.Show("Calculated length exceeds maximum of " & MaxLength, "Warning")
        newLength = MaxLength
    End If
    
    ' Update parameter
    Parameter("Length") = newLength
    
    ' Update model
    iLogicVb.UpdateWhenDone = True
    
    ' Success message
    MessageBox.Show("Length updated to " & newLength & " mm", "Success")
    
Catch ex As Exception
    MessageBox.Show("Error: " & ex.Message, "iLogic Error")
End Try
```

## Next Steps

- Learn about [Parameters and Properties](./02-parameters-properties.md)
- Explore [Rules and External Rules](./03-rules-external-rules.md)
- See [Common Patterns](../common-patterns/) for practical examples
