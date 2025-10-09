' Title: is there a way to find all Browser notes that does not show the partnumber ?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/is-there-a-way-to-find-all-browser-notes-that-does-not-show-the/td-p/13815636#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:56:03.523707

'this my current code

'Reset Browsernote that dont show the partnummer
Sub Main()
	
	Dim doc = ThisApplication.ActiveDocument
	Dim sDocumentSubType As String = doc.SubType
	
	Try
		'debug("Start ilogic: RenameBrowsernoteAll")
		Logger.Info("Start ilogic: RenameBrowsernoteAll")
		
		oDoc = ThisDoc.Document
		Dim oPane As BrowserPane
		Dim oTopNode As BrowserNode
		
			If oDoc.DocumentType = kPartDocumentObject
				oPane = oDoc.BrowserPanes.Item("PmDefault")
			Else 'kAssemblyDocumentObject or 
					oPane = oDoc.BrowserPanes.Item("Model")
			End If
			
				oTopNode = oPane.TopNode
			
			If oTopNode.BrowserNodeDefinition.Label <> System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName) Then
				oDoc.DisplayName = "" 'System.IO.Path.GetFileNameWithoutExtension(oDoc.FullFileName)
			End If 
	
	Catch
		debug("RenameBrowsernoteALL ilogic code fail for unknown reason")
	End Try
If sDocumentSubType = "{E60F81E1-49B3-11D0-93C3-7E0706000000}" Or sDocumentSubType = "{28EC8354-9024-440F-A8A2-0E0E55D635B0}" Then ' = "assembly"''

If(doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = iProperties.Value("Project", "Part Number")) Then

Assembly

Else 
	
	Logger.Info("This is not the Aktive Doc - End Code")
	
Exit Sub


End If

End If

'debug("End ilogic: RenameBrowsernoteAll")
Logger.Info("End ilogic: RenameBrowsernoteAll")

End Sub

Public Sub debug(txt As String)
        Trace.WriteLine("NTI : " & txt)
End Sub


Sub Assembly()
	Logger.Info("Change Browsernotes Assembly", filename)
	
	Dim oAsmDoc As AssemblyDocument
	oAsmDoc = ThisApplication.ActiveDocument

	If (ThisDoc.Document.DisplayName = oAsmDoc.DisplayName) Then
	
	
	Else
		Logger.Info("Don´t run - End Code", filename)

		Exit Sub
	End If
	

Dim oRefDocs As DocumentsEnumerator
oRefDocs = oAsmDoc.AllReferencedDocuments
Dim oRefDoc As Document
 
For Each oRefDoc In oRefDocs
	
	Try
	
	
	Catch
		
		
		Try
			
			If oRefDoc.DisplayName <> oRefDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value Then
				
				Logger.Info("All need To be true " & (Left(oRefDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value,9) <> "Zero Line").ToString & " " &  (Left(oRefDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value,9) <> "Zero line").ToString & " " & (Left(oRefDoc.PropertySets("Design Tracking Properties").Item("Part Number").Value,15) <> "Control cabinet").ToString, filename)
	
					If Left(oRefDoc.DisplayName,9) <> "Zero Line" And Left(oRefDoc.DisplayName,9) <> "Zero line" And Left(oRefDoc.DisplayName,15) <> "Control cabinet"   Then
					oRefDoc.DisplayName = "" 
					End If
	
				
			End If 
		
		Catch
			
		End Try
	
	End Try

Next
Dim asm As AssemblyDocument = ThisDoc.Document
ResetOccurrenceNames(asm.ComponentDefinition.Occurrences)

End Sub

Sub ResetOccurrenceNames(occurrences As Object)
	
	Try
	For Each occ As ComponentOccurrence In occurrences
		Dim dispName = occ.ReferencedDocumentDescriptor.DisplayName
		
		'start
		Dim oPartDoc As Document
    oPartDoc = occ.Definition.Document
    Dim partNumber As String
    partNumber = iProperties.Value(oPartDoc, "Project", "Part Number")
		'end
		
		
		Dim length = Len(partNumber) -4
		
		If Not occ.Name.StartsWith(Left(dispName,length)) Then
		Logger.Info(occ.Name.StartsWith(Left(dispName,length)), "2")
If(Left(occ.Name, 9)) <> "Zero Line" And (Left(occ.Name, 9)) <> "Zero line" And (Left(occ.Name, 15)) <> "Control cabinet" Then
		occ.Name = ""
	End If 
		Logger.Info("Reset browsernote", filename)
		
		Else
			If Left(occ.Name.ToString, Len(dispName)) = dispName Then
			
			Else
			Logger.Info(occ.Name.StartsWith(Left(dispName,length)) & " 2" , "2")

			Logger.Info("Browser Note " & occ.Name, filename)
If (Left(occ.Name, 9)) <> "Zero Line" And (Left(occ.Name, 9)) <> "Zero line" And (Left(occ.Name, 15)) <> "Control cabinet" Then
			occ.Name = ""
			
		End If 
		
		End If
	End If 
		'ResetOccurrenceNames(occ.SubOccurrences)
	Next
	Catch
		Logger.Info("somthing fail "   , filename)

	End Try
End Sub