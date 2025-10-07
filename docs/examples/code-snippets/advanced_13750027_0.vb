' Title: I need with a rule that will look for the closed profiles in a sketch array and select them all
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/i-need-with-a-rule-that-will-look-for-the-closed-profiles-in-a/td-p/13750027#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:08:51.691353

' Get the active part document
Dim oPartDoc As PartDocument = ThisApplication.ActiveDocument

' Get the part component definition
Dim oPartDef As PartComponentDefinition = oPartDoc.ComponentDefinition

' Find the specific sketch by name
Dim oSketch As PlanarSketch = Nothing
For i = 1 To oPartDef.Sketches.Count
    If oPartDef.Sketches.Item(i).Name = "Bottom Hoop Profiles" Then
        oSketch = oPartDef.Sketches.Item(i)
        Exit For
    End If
Next

' Check if sketch was found
If oSketch Is Nothing Then
    MessageBox.Show("Sketch 'Bottom Hoop Profiles' not found!", "Error")
    Exit Sub
End If

' For sketch arrays, we need to look at the sketch entities and patterns
MessageBox.Show("Found sketch: " & oSketch.Name & vbCrLf & "Profile count: " & oSketch.Profiles.Count, "Debug Info")

' Method 1: Try to work with all profiles (including arrayed ones)
If oSketch.Profiles.Count > 0 Then
    
    ' Loop through each profile in the sketch
    For i = 1 To oSketch.Profiles.Count
        Dim oProfile As Profile = oSketch.Profiles.Item(i)
        
        Try
            ' For arrayed profiles, we might need to check differently
            Dim profilePaths As Integer = oProfile.Count
            MessageBox.Show("Profile " & i & " has " & profilePaths & " path(s)", "Debug")
            
            ' Try to revolve each profile
            Try
                ' Create revolve definition around Y-axis (WorkAxis 2)
                Dim oRevolveDef = oPartDef.Features.RevolveFeatures.CreateRevolveDefinition(oProfile, oPartDef.WorkAxes.Item(2))
                
                ' Set revolve to full 360 degrees
                oRevolveDef.SetToFullSweep()
                
                ' Create the revolve feature
                Dim oRevolveFeature = oPartDef.Features.RevolveFeatures.Add(oRevolveDef)
                
                MessageBox.Show("Successfully revolved profile " & i, "Success")
                
            Catch ex As Exception
                ' Handle any errors (e.g., profile can't be revolved)
                MessageBox.Show("Could not revolve profile " & i & ": " & ex.Message, "Revolve Error")
            End Try
            
        Catch ex As Exception
            MessageBox.Show("Error processing profile " & i & ": " & ex.Message, "Error")
        End Try
        
    Next
    
Else
    MessageBox.Show("No profiles found. This might be due to how sketch arrays are handled.", "Info")
End If

' Alternative Method 2: Look for sketch patterns and work with them
Try
    If oSketch.SketchPatterns.Count > 0 Then
        MessageBox.Show("Found " & oSketch.SketchPatterns.Count & " sketch patterns", "Pattern Info")
        
        ' You might need to suppress the pattern temporarily and work with individual profiles
        ' Or create a separate approach for patterned geometry
    End If
Catch
    ' SketchPatterns might not be available in this context
End Try

' Update the part
oPartDoc.Update()

MessageBox.Show("Operation completed!", "Final")