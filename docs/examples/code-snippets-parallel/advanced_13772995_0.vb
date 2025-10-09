' Title: Batch Updated Drawing Number
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/batch-updated-drawing-number/td-p/13772995
' Category: advanced
' Scraped: 2025-10-09T09:03:51.327292

Dim oDrawDoc As DrawingDocument
Dim oSheet As Sheet
Dim tb As Inventor.TextBox
Dim oBorder As Border
Dim borderDef As BorderDefinition
Dim oTB1'oTB1 references to Titleblocks
Dim sValue As String
Dim oDrwNo 
Dim oPromptEnt As String

oDrawDoc = ThisApplication.ActiveDocument
oSheet = oDrawDoc.ActiveSheet
oBorder = oSheet.Border
borderDef = oBorder.Definition
oTB1 = oSheet.TitleBlock

On Error Resume Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "DRAWING NO." Then
						sValue = oTB1.GetResultText(tb)
					DRAWING_NO = sValue
'					MessageBox.Show("DWG NO " & DRAWING_NO, "Title")

						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
						'MessageBox.Show("Param drw no  " & oDrwNo, "Title")
			'Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oDrwNo) 
						'oPromptEnt = oTB1.GetResultText(tb)
						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
		

	End If
Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "<TITLE 3>" Then
						sValue = oTB1.GetResultText(tb)
						TITLE_LINE_3 = sValue 
						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
'						'MessageBox.Show("Param drw no  " & oTitle3, "Title")
'			Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oTitle3) 
'						oPromptEnt = oTB1.GetResultText(tb)
'						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
'		
'
	End If
Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "<REVISION NUMBER>" Then
						sValue = oTB1.GetResultText(tb)
						SHEET_REVISION_NO = sValue
'						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
'						'MessageBox.Show("Param drw no  " & oShtRevNo, "Title")
'			Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oShtRevNo) 
'						oPromptEnt = oTB1.GetResultText(tb)
'						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
'		
'
	End If
Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "REV 1" Then
						sValue = oTB1.GetResultText(tb)
						FIRST_REV_NO = sValue
						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
'						'MessageBox.Show("Param drw no  " & oRev1, "Title")
'			Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oRev1) 
'						oPromptEnt = oTB1.GetResultText(tb)
'						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
'		
'
	End If
Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "REV 1 DESCRIPTION" Then
						sValue = oTB1.GetResultText(tb)
						FIRST_REV_DESCRIPTION = sValue
'						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
'						'MessageBox.Show("Param drw no  " & oRevDesc, "Title")
'			Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oRevDesc) 
'						oPromptEnt = oTB1.GetResultText(tb)
'						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
'		
'
	End If
Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "REV 1 BY" Then
						sValue = oTB1.GetResultText(tb)
						FIRST_REV_BY = sValue
						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
'						'MessageBox.Show("Param drw no  " & oRev1By, "Title")
'			Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oRev1By) 
'						oPromptEnt = oTB1.GetResultText(tb)
'						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
'		
'
	End If
Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "REV 1 APPROVED BY" Then
						sValue = oTB1.GetResultText(tb)
						FIRST_REV_APPROVED_BY = sValue
						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
'						'MessageBox.Show("Param drw no  " & oRev1AppBy, "Title")
'			Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oRev1AppBy) 
'						oPromptEnt = oTB1.GetResultText(tb)
'						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
'		
'
	End If
Next

For Each tb In oDrawDoc.ActiveSheet.TitleBlock.Definition.Sketch.TextBoxes 
If tb.Text = "REV 1 DATE" Then
						sValue = oTB1.GetResultText(tb)
						FIRST_REV_DATE = sValue
						'MessageBox.Show("Property Field: " & tb.Text & vbCrLf & "Value: " & sValue & vbCrLf & "Prompted Entry: " & tb.FormattedText, "Prompted entry")        
'						'MessageBox.Show("Param drw no  " & oRev1Date, "Title")
'			Call oDrawDoc.ActiveSheet.TitleBlock.SetPromptResultText(tb, oRev1Date) 
'						oPromptEnt = oTB1.GetResultText(tb)
'						'MessageBox.Show("Prompted value " & oPromptEnt, "Title")
'		
'
	End If
Next