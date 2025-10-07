' Title: iLogic Rule: Rename Browser Nodes to Show Part Number + Title
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-rename-browser-nodes-to-show-part-number-title/td-p/13836335
' Category: advanced
' Scraped: 2025-10-07T13:59:51.023769

' iLogic code to change part display names in assembly to Part Number + Title
' This code should be run from within the assembly document

Sub Main()
    ' Get the active document (should be an assembly)
    Dim oAssemblyDoc As AssemblyDocument
    Try
        oAssemblyDoc = ThisApplication.ActiveDocument
    Catch
        MessageBox.Show("Please run this code from an assembly document", "Error")
        Exit Sub
    End Try
    
    ' Check if it's actually an assembly
    If oAssemblyDoc.DocumentType <> kAssemblyDocumentObject Then
        MessageBox.Show("This code must be run from an assembly document", "Error")
        Exit Sub
    End If
    
    ' Get the assembly definition
    Dim oAssemblyDef As AssemblyComponentDefinition
    oAssemblyDef = oAssemblyDoc.ComponentDefinition
    
    ' Counter for processed components
    Dim processedCount As Integer = 0
    Dim skippedCount As Integer = 0
    Dim debugInfo As String = ""
    
    ' Loop through all occurrences in the assembly
    For Each oOccurrence As ComponentOccurrence In oAssemblyDef.Occurrences
        Try
            ' Skip Content Center parts - check if it's from Content Center
            Dim isContentCenter As Boolean = False
            Try
                ' Check file path for Content Center
                Dim filePath As String = oOccurrence.Definition.Document.FullFileName
                If filePath.ToUpper().Contains("CONTENT CENTER") Or _
                   filePath.ToUpper().Contains("CONTENTCENTER") Or _
                   filePath.ToUpper().Contains("LIBRARIES") Then
                    isContentCenter = True
                End If
            Catch
                ' Alternative method - check if it's a library component
                Try
                    If oOccurrence.IsLibraryComponent Then
                        isContentCenter = True
                    End If
                Catch
                    ' If we can't determine, assume it's not Content Center
                    isContentCenter = False
                End Try
            End Try
            
            If isContentCenter Then
                skippedCount = skippedCount + 1
                debugInfo = debugInfo & "Skipped Content Center: " & oOccurrence.Name & vbCrLf
                Continue For
            End If
            
            ' Get the part document from the occurrence
            Dim oPartDoc As Document = oOccurrence.Definition.Document
            
            ' Get part number and title from iProperties
            Dim partNumber As String = ""
            Dim partTitle As String = ""
            
            ' Try multiple ways to get Part Number
            Try
                ' First try Design Tracking Properties
                partNumber = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value.ToString()
            Catch
                Try
                    ' Try Summary Information
                    partNumber = oPartDoc.PropertySets.Item("Summary Information").Item("Title").Value.ToString()
                Catch
                    Try
                        ' Try Inventor Summary Information
                        partNumber = oPartDoc.PropertySets.Item("Inventor Summary Information").Item("Part Number").Value.ToString()
                    Catch
                        ' Use filename without extension as fallback
                        partNumber = System.IO.Path.GetFileNameWithoutExtension(oPartDoc.FullFileName)
                    End Try
                End Try
            End Try
            
            ' Try multiple ways to get Title/Description
            Try
                ' First try Design Tracking Properties
                partTitle = oPartDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value.ToString()
            Catch
                Try
                    ' Try Summary Information
                    partTitle = oPartDoc.PropertySets.Item("Summary Information").Item("Title").Value.ToString()
                Catch
                    Try
                        ' Try Inventor Summary Information
                        partTitle = oPartDoc.PropertySets.Item("Inventor Summary Information").Item("Title").Value.ToString()
                    Catch
                        partTitle = "No Description"
                    End Try
                End Try
            End Try
            
            ' Clean up empty values
            If String.IsNullOrEmpty(partNumber) Then partNumber = "No Part Number"
            If String.IsNullOrEmpty(partTitle) Then partTitle = "No Description"
            
            ' Create new display name, preserving occurrence suffix (like :1, :2, etc.)
            Dim newDisplayName As String = partNumber & " - " & partTitle
            
            ' Check if the original name has an occurrence suffix (like :1, :2, etc.)
            Dim oldName As String = oOccurrence.Name
            Dim occurrenceSuffix As String = ""
            
            ' Look for pattern like ":1", ":2", etc. at the end of the name
            Dim colonIndex As Integer = oldName.LastIndexOf(":")
            If colonIndex > 0 Then
                Dim possibleSuffix As String = oldName.Substring(colonIndex)
                ' Check if it's a number after the colon
                Dim suffixNumber As Integer
                If Integer.TryParse(possibleSuffix.Substring(1), suffixNumber) Then
                    occurrenceSuffix = possibleSuffix
                End If
            End If
            
            ' Add the occurrence suffix back if it existed
            If occurrenceSuffix <> "" Then
                newDisplayName = newDisplayName & occurrenceSuffix
            End If
            
            ' Set the new display name for the occurrence
            oOccurrence.Name = newDisplayName
            
            processedCount = processedCount + 1
            debugInfo = debugInfo & "Updated: " & oldName & " → " & newDisplayName & vbCrLf
            
        Catch ex As Exception
            ' Continue processing other parts if one fails
            debugInfo = debugInfo & "Error processing: " & oOccurrence.Name & " - " & ex.Message & vbCrLf
        End Try
    Next
    
    ' Update the assembly
    oAssemblyDoc.Update()
    
    ' Show simple completion message
    Dim message As String = "Part names updated successfully!" & vbCrLf & vbCrLf & _
                           "Processed: " & processedCount.ToString() & " components" & vbCrLf & _
                           "Skipped: " & skippedCount.ToString() & " Content Center parts"
    
    MessageBox.Show(message, "Complete")
    
End Sub