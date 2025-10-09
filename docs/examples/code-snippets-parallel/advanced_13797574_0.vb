' Title: How To Delete Occurrence Pattern Element
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-delete-occurrence-pattern-element/td-p/13797574
' Category: advanced
' Scraped: 2025-10-09T08:56:43.708945

Option Explicit

Sub SuppressPatternOccurrencesByName()

    Dim asmDoc As AssemblyDocument
    Dim asmDef As AssemblyComponentDefinition
    Dim occPattern As OccurrencePattern
    Dim targetName As String
    Dim i As Integer
    
    Dim rectPattern As RectangularOccurrencePattern
    Dim circPattern As CircularOccurrencePattern
    Dim compOcc As ComponentOccurrence
    Dim occPatternElem As OccurrencePatternElement
    
    targetName = "OIL OUTLET"
    
    Set asmDoc = ThisApplication.ActiveDocument
    If asmDoc.DocumentType <> kAssemblyDocumentObject Then
        MsgBox "Please open an Assembly file before running this macro!", vbExclamation
        Exit Sub
    End If
    
    Set asmDef = asmDoc.ComponentDefinition
    
    For Each occPattern In asmDef.OccurrencePatterns
        
        If TypeOf occPattern Is RectangularOccurrencePattern Then
            Set rectPattern = occPattern
            
            For Each occPatternElem In rectPattern.OccurrencePatternElements
                For i = occPatternElem.occurrences.count To 1 Step -1
                    If InStr(1, occPatternElem.occurrences.item(i).name, targetName, vbTextCompare) <> 0 Then
                        occPatternElem.occurrences.item(i).Delete2 (False)
                    End If
                Next i
            Next occPatternElem

        End If
        
    Next occPattern

End Sub