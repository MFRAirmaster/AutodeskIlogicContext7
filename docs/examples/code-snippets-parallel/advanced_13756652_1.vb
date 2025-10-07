' Title: Automated Mates
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automated-mates/td-p/13756652#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:08:09.283912

Dim oAsm As AssemblyDocument = ThisApplication.ActiveDocument
Dim oDef As AssemblyComponentDefinition = oAsm.ComponentDefinition

Dim asmXY As WorkPlane = oDef.WorkPlanes.Item("XY Plane")
Dim asmYZ As WorkPlane = oDef.WorkPlanes.Item("YZ Plane")
Dim asmXZ As WorkPlane = oDef.WorkPlanes.Item("XZ Plane")

For Each oOcc As ComponentOccurrence In oDef.Occurrences
    If oOcc.Grounded Then Continue For

    ' Optional: Temporarily skip known problematic components
    ' If oOcc.Name = "P25237-CMS-QTS-QTS-RIC2:1" Or oOcc.Name = "25237-WALL-1" Then Continue For

    Dim compDef As Object = oOcc.Definition
    Dim xyPlane As WorkPlane = Nothing
    Dim yzPlane As WorkPlane = Nothing
    Dim xzPlane As WorkPlane = Nothing

    ' Check component type and retrieve work planes
    If TypeOf compDef Is PartComponentDefinition Then
        Dim partDef As PartComponentDefinition = compDef
        xyPlane = TryCast(partDef.WorkPlanes.Item("XY Plane"), WorkPlane)
        yzPlane = TryCast(partDef.WorkPlanes.Item("YZ Plane"), WorkPlane)
        xzPlane = TryCast(partDef.WorkPlanes.Item("XZ Plane"), WorkPlane)
    ElseIf TypeOf compDef Is AssemblyComponentDefinition Then
        Dim asmDef As AssemblyComponentDefinition = compDef
        xyPlane = TryCast(asmDef.WorkPlanes.Item("XY Plane"), WorkPlane)
        yzPlane = TryCast(asmDef.WorkPlanes.Item("YZ Plane"), WorkPlane)
        xzPlane = TryCast(asmDef.WorkPlanes.Item("XZ Plane"), WorkPlane)
    Else
        MessageBox.Show("Unsupported component type: " & oOcc.Name)
        Continue For
    End If

    ' Skip if any plane is missing
    If xyPlane Is Nothing Or yzPlane Is Nothing Or xzPlane Is Nothing Then
        MessageBox.Show("Missing work plane for: " & oOcc.Name)
        Continue For
    End If

    ' Check component state
    If oOcc.Suppressed Then
        MessageBox.Show("Component suppressed: " & oOcc.Name)
        Continue For
    End If
    If oOcc.IsPatternElement Then
        MessageBox.Show("Component in pattern: " & oOcc.Name)
        Continue For
    End If

    ' Apply flush constraints
    Try
        oDef.Constraints.AddFlushConstraint(asmXY, xyPlane, 0)
        oDef.Constraints.AddFlushConstraint(asmYZ, yzPlane, 0)
        oDef.Constraints.AddFlushConstraint(asmXZ, xzPlane, 0)
        oOcc.Grounded = False
    Catch ex As Exception
        Dim compType As String = If(TypeOf compDef Is PartComponentDefinition, "Part", "Assembly")
        MessageBox.Show("Constraint error with: " & oOcc.Name & vbCrLf & _
                        "Type: " & compType & vbCrLf & _
                        "Details: " & ex.Message & vbCrLf & _
                        "Source: " & ex.Source)
        Continue For
    End Try
Next