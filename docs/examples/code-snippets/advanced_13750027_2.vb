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

Try
	Dim newProfile As Profile
	newProfile = oSketch.Profiles.AddForSolid
    Try
        ' Create revolve definition around Y-axis (WorkAxis 2)
        Dim oRevolveDef = oPartDef.Features.RevolveFeatures.AddFull(newProfile, oPartDef.WorkAxes.Item(2), PartFeatureOperationEnum.kNewBodyOperation)    
    Catch ex As Exception
        ' Handle any errors (e.g., profile can't be revolved)
        MessageBox.Show("Could not revolve profile: "  & ex.Message, "Revolve Error")
    End Try
Catch
    MessageBox.Show("Could not extract profiles", "Info")
End Try