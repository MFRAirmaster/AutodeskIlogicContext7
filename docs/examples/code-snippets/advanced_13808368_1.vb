' Title: ilogic code needed: Finish Feature, is it used, if so set the first one to an iproperty
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-needed-finish-feature-is-it-used-if-so-set-the-first/td-p/13808368#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:13:28.061042

Dim finnishft As FinishFeature
For Each ft As PartFeature In ThisDoc.Document.componentdefinition.features
If ft.Type = ObjectTypeEnum.kFinishFeatureObject Then
finnishft = ft
Exit For
End If
Next
Dim finisharea As Double = 0
Try
finisharea=	finnishft.Faces(1).Evaluator.Area
Catch ex As Exception
	
End Try
MsgBox (finisharea)