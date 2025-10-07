' Title: update drawing line colour based on part colour
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/update-drawing-line-colour-based-on-part-colour/td-p/7417136
' Category: advanced
' Scraped: 2025-10-07T13:59:00.967123

Sub Main 
	' Get the active drawing document. 
	Dim oDoc As DrawingDocument 
	oDoc = ThisApplication.ActiveDocument 
	
	Dim oSheets As Sheets
	oSheets = oDoc.Sheets
	Dim oSheet As Sheet
	
	'get current sheet so it can
	'be made active again later
	Dim oCurrentSheet As Sheet
	oCurrentSheet = oDoc.ActiveSheet	
	
	Dim oViews As DrawingViews
	Dim oView As DrawingView
	
	' Iterate through the sheets
	For Each oSheet In oSheets
	'	activate the sheet
		oSheet.Activate
		
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
			
			' Process the view, wrapping it in a transaction so the 
			' each view can be undone with a single undo operation. 
			Dim trans As Transaction 
			trans = ThisApplication.TransactionManager.StartTransaction( _ 
									oDoc, "Change drawing view color")  
			
			' Call the recursive function that does all the work. 
			Call ProcessAssemblyColor(oView, asmDef.Occurrences) 
			trans.End 
		Next
		'update the sheet
		oSheet.Update
	Next
	
	'return to original sheet	
	oCurrentSheet.Activate
End Sub 



Private Sub ProcessAssemblyColor(drawView As DrawingView, _ 
                                 Occurrences As ComponentOccurrences) 
   ' Iterate through the current collection of occurrences. 
   Dim occ As ComponentOccurrence 
   For Each occ In Occurrences 
      ' Check to see if this occurrence is a part or assembly. 
      If occ.DefinitionDocumentType = kPartDocumentObject Then 
         ' ** It's a part so process the color.  

         ' Get the render style of the occurrence. 
         Dim color As RenderStyle 
         Dim sourceType As StyleSourceTypeEnum 
         color = occ.GetRenderStyle(sourceType)  

         ' Get the TransientsObjects object to use later. 
         Dim transObjs As TransientObjects 
         transObjs = ThisApplication.TransientObjects  

         ' Verify that a layer exists for this color. 
         Dim layers As LayersEnumerator 
         layers = drawView.Parent.Parent.StylesManager.layers  

         Dim oDoc As DrawingDocument 
         oDoc = drawView.Parent.Parent  

         On Error Resume Next 
         Dim colorLayer As Layer 
         colorLayer = layers.Item(color.Name)  

         If Err.Number <> 0 Then 
            On Error Goto 0 
            ' Get the diffuse color for the render style. 
            Dim red As Byte 
            Dim green As Byte 
            Dim blue As Byte  

            ' Create a color object that is the diffuse color. 
            Call color.GetDiffuseColor(red, green, blue) 
            Dim newColor As color 
            newColor = transObjs.CreateColor(red, green, blue)  

            ' Copy an arbitrary layer giving it the name 
            ' of the render style. 
            colorLayer = layers.Item(1).Copy(color.Name) 

            ' the attributes of the layer to use the color, 
            ' have a solid line type, and a specific width. 
            colorLayer.Color = newColor 
            colorLayer.LineType = kContinuousLineType 
            colorLayer.LineWeight = 0.02 
         End If 
         On Error Goto 0  

         ' Get all of the curves associated with this occurrence. 
         On Error Resume Next 
         Dim drawcurves As DrawingCurvesEnumerator 
         drawcurves = drawView.DrawingCurves(occ) 
         If Err.Number = 0 Then 
            On Error Goto 0  

            ' Create an empty collection. 
            Dim objColl As ObjectCollection 
            objColl = transObjs.CreateObjectCollection()  

            ' Add the curve segments to the collection. 
            Dim drawCurve As DrawingCurve 
            For Each drawCurve In drawcurves 
               Dim segment As DrawingCurveSegment 
               For Each segment In drawCurve.Segments 
                  objColl.Add (segment)
               Next 
            Next  

            ' Change the layer of all of the segments. 
            Call drawView.Parent.ChangeLayer(objColl, colorLayer) 
         End If 
         On Error Goto 0 
      Else 
         ' It's an assembly so process its contents. 
         Call ProcessAssemblyColor(drawView, occ.SubOccurrences) 
      End If 
   Next 
End Sub