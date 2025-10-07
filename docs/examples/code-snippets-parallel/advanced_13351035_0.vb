' Title: ID of occurrence in a pattern
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/id-of-occurrence-in-a-pattern/td-p/13351035#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:37:53.038977

Sub Main
	Dim oPDoc As PartDocument = TryCast(ThisDoc.Document, Inventor.PartDocument)
	If oPDoc Is Nothing Then Return
	Dim oRPFs As RectangularPatternFeatures = oPDoc.ComponentDefinition.Features.RectangularPatternFeatures
	Dim oMyRPF As RectangularPatternFeature = Nothing
	For Each oRPF As RectangularPatternFeature In oRPFs
		If oRPF.Name = "Rectangular Pattern1" Then
			oMyRPF = oRPF
			Exit For 'exit iteration once found
		End If
	Next
	If oMyRPF Is Nothing Then
		MsgBox("Pattern Not Found!", vbCritical, "iLogic")
		Return
	End If
	Dim oMyFPE As FeaturePatternElement = Nothing
	For Each oFPE As FeaturePatternElement In oMyRPF.PatternElements
		If oFPE.Index = 8 Then
			oMyFPE = oFPE
			Exit For 'exit iteration once found
		End If
	Next
	If oMyFPE Is Nothing Then
		MsgBox("Pattern Element At That Index Not Found!", vbCritical, "iLogic")
		Return
	End If
	'Toggle its suppressed status
	oMyFPE.Suppressed = Not oMyFPE.Suppressed
	oPDoc.Update()
End Sub