' Title: update drawing line colour based on part colour
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-drawing-line-colour-based-on-part-colour/td-p/7417136
' Category: advanced
' Scraped: 2025-10-09T09:03:39.887945

Sub Main 
	' Get the active drawing document. 
	Dim oDoc As DrawingDocument 
	oDoc = ThisApplication.ActiveDocument 
	
	Dim oSheet As Sheet
	oSheet = oDoc.ActiveSheet	
	
	Dim oViews As DrawingViews
	Dim oView As DrawingView
	

	'get the collection of view on the sheet
	oViews = oSheet.DrawingViews		
	
	' Iterate through the views on the sheet
	For Each oView In oViews	
		
		Dim docDesc As DocumentDescriptor 
		docDesc = oView.ReferencedDocumentDescriptor  
		
		' Verify that the drawing view is of an assembly. 
		If docDesc.ReferencedDocumentType <> kAssemblyDocumentObject Then 
			Continue For
		End If  
		
		' Get the component definition for the assembly. 
		Dim asmDef As AssemblyComponentDefinition 
		asmDef = docDesc.ReferencedDocument.ComponentDefinition 
		
		'define view rep collection
		Dim oViewReps As DesignViewRepresentations
		oViewReps = asmDef.RepresentationsManager.DesignViewRepresentations
		
		oColourViewRepFound = False
		For Each oViewRep In oViewReps
			If oViewRep.Name = "Colour" Then
				oColourViewRepFound = True 
			End If
		Next
		
		If oColourViewRepFound = True Then
			Call ProcessAssemblyColor(oView, asmDef.Occurrences) 
		End If
	Next

End Sub 



Private Sub ProcessAssemblyColor(drawView As DrawingView, _ 
                                 Occurrences As ComponentOccurrences) 
   ' Iterate through the current collection of occurrences. 
   Dim occ As ComponentOccurrence 
   For Each occ In Occurrences 
      ' Check to see if this occurrence is a part or assembly. 
      If occ.DefinitionDocumentType = kPartDocumentObject Then 

         ' Get the render style of the occurrence. 
         Dim color As RenderStyle 
         Dim sourceType As StyleSourceTypeEnum 
         color = occ.GetRenderStyle(sourceType)  
		If color.name IsNot Nothing Then
		 iProperties.Value("Custom", "Colour") = color.name
		 Exit For
	 	End If
			

      End If 
   Next 
End Sub