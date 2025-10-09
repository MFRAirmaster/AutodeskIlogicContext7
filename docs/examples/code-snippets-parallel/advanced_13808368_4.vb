' Title: ilogic code needed: Finish Feature, is it used, if so set the first one to an iproperty
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-needed-finish-feature-is-it-used-if-so-set-the-first/td-p/13808368#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:48:04.414652

'USING THE FINISH FEATURE.
'USES THE FINISH SHORT DESCRIPTION FOR THE FINISH IPROPERTY, WHICH IS USED ON THE DRAWING BORDER
'LOOKS AT THE AREA Of THE FINISH AREA VS THE AREA Of THE PART And ADDS "EXCEPT AS NOTED" To THE FINISH NOTE
Dim oDoc As Document = ThisApplication.ActiveDocument
Dim hasFinishFeature As Boolean = False
Dim finishft As FinishFeature
For Each ft As PartFeature In ThisDoc.Document.componentdefinition.features
	If ft.Type = ObjectTypeEnum.kFinishFeatureObject Then
		finishft = ft
		hasFinishFeature = True
		Exit For
	End If
	Next

If hasFinishFeature = False Then
	MessageBox.Show("No Finishes Exist", "Finish Error")
	iProperties.Value("Custom", "Finish")=""
	Return
Else
	Dim finisharea As Double = 0
	Try
	finisharea = finishft.Faces(1).Evaluator.Area * 0.0393701
	'0.0393701 converts Area in mm^2 to in^2.
	Catch ex As Exception
	End Try
	surfaceArea = iProperties.Area
	If finisharea = 0 Then
		FinishComment = ""
	ElseIf surfaceArea - finisharea >.005 Then 
		FinishComment = ", EXCEPT AS NOTED"
	End If
		
	Dim oFinishFeature As FinishFeature = oDoc.ComponentDefinition.Features.FinishFeatures.Item(1)
	Dim oFinishDef As FinishDefinition = oFinishFeature.Definition
	Dim myDescription As String = ""
	Try
	myDescription = oFinishDef.ShortDescription
	Catch ex As Exception
	End Try
	iProperties.Value("Custom", "Finish") = myDescription + FinishComment
End If