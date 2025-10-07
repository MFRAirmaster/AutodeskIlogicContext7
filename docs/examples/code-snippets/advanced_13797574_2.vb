' Title: How To Delete Occurrence Pattern Element
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-delete-occurrence-pattern-element/td-p/13797574
' Category: advanced
' Scraped: 2025-10-07T12:48:14.621391

Sub main()

    Dim asmDoc As AssemblyDocument = ThisApplication.ActiveDocument
    Dim asmDef As AssemblyComponentDefinition = asmDoc.ComponentDefinition
    Dim occPattern As OccurrencePattern
    Dim targetName As String
    
    Dim rectPattern As RectangularOccurrencePattern
    Dim circPattern As CircularOccurrencePattern
    Dim compOcc As ComponentOccurrence
    Dim occPatternElem As OccurrencePatternElement
    
    targetName = "OIL OUTLET"
    
    If asmDoc.DocumentType <> kAssemblyDocumentObject Then
        MsgBox("Please open an Assembly file before running this macro!", vbExclamation)
        Exit Sub
    End If
	
    For Each occPattern In asmDef.OccurrencePatterns
        
        If TypeOf occPattern Is RectangularOccurrencePattern Then
            rectPattern = occPattern
			Dim i As Integer = 0
			For k As Integer = 1 To rectPattern.ParentComponents.Count
				If rectPattern.ParentComponents(k).Name.Contains(targetName) Then
					i = k
					Exit For
				End If
			Next k
			If i = 0 Then Continue For
			Dim oObjs As ObjectCollection
			oObjs = ThisApplication.TransientObjects.CreateObjectCollection(rectPattern.ParentComponents)
            oObjs.Remove(i)
			rectPattern.ParentComponents = oObjs

        End If
        
    Next occPattern

End Sub