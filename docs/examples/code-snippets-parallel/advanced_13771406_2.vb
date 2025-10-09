' Title: ilogic Rule in part runs from assembly.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-in-part-runs-from-assembly/td-p/13771406#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:49:15.393416

Dim oPartDoc As PartDocument
oPartDoc = ThisApplication.ActiveEditDocument
Dim oCompDef As ComponentDefinition
oCompDef = oPartDoc.ComponentDefinition



Dim oFaces As Faces
Dim oFace As Face
Dim oSurfBodies As SurfaceBodies
Dim oSurfBody As SurfaceBody
Dim oStyle As RenderStyle
	oStyle = oPartDoc.RenderStyles.Item("Green") 
	 oSurfBodies = oCompDef.SurfaceBodies

oFaces = oSurfBodies.Item(1).Faces
For Each oSurfBody In oSurfBodies
oFaces = oSurfBody.Faces
	For Each oFace In oFaces
		For i = 1 To 2
			Try			
			If oFace.AttributeSets.Item(i).Item(1).Value = "Long1" Then
				If iProperties.Value("Custom", "Long1") = "1" Or iProperties.Value("Custom", "Long1") = "2" Then
				Call oFace.SetRenderStyle(kOverrideRenderStyle, oStyle)
				Else
				Call oFace.SetRenderStyle(kPartRenderStyle)
				End If 
			End If
			
			If oFace.AttributeSets.Item(i).Item(1).Value = "Long2" Then
				If iProperties.Value("Custom", "Long2") = "1" Or iProperties.Value("Custom", "Long2") = "2" Then
				Call oFace.SetRenderStyle(kOverrideRenderStyle, oStyle)
				Else
				Call oFace.SetRenderStyle(kPartRenderStyle)
				End If 
			End If
			
			If oFace.AttributeSets.Item(i).Item(1).Value = "Short1" Then
				If iProperties.Value("Custom", "Short1") = "1" Or iProperties.Value("Custom", "Short1") = "2" Then
				Call oFace.SetRenderStyle(kOverrideRenderStyle, oStyle)
				Else
				Call oFace.SetRenderStyle(kPartRenderStyle)
				End If 
			End If
			
			If oFace.AttributeSets.Item(i).Item(1).Value = "Skort2" Then
				If iProperties.Value("Custom", "Short2") = "1" Or iProperties.Value("Custom", "Short2") = "2" Then
				Call oFace.SetRenderStyle(kOverrideRenderStyle, oStyle)
				Else
				Call oFace.SetRenderStyle(kPartRenderStyle)
				End If 
			End If
			Catch	
			End Try
		Next		
	Next
Next