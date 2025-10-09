' Title: Converting an Inventor .ipt model into Ghost / Reference Geometry using ClientGraphics.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/converting-an-inventor-ipt-model-into-ghost-reference-geometry/td-p/13842968#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:08:01.998946

Sub Main
	Dim oInvApp As Inventor.Application = ThisApplication
	Dim oPDoc As PartDocument = TryCast(oInvApp.ActiveDocument, Inventor.PartDocument)
	If oPDoc Is Nothing Then Return
	Dim oGDSC As Inventor.GraphicsDataSetsCollection
	Try : oGDSC = oPDoc.GraphicsDataSetsCollection : Catch : End Try
	If oGDSC IsNot Nothing AndAlso oGDSC.Count > 0 Then
		For Each oGDS As Inventor.GraphicsDataSets In oGDSC
			Try
				oGDS.Delete()
			Catch ex As Exception
				Logger.Error(ex.ToString())
			End Try
		Next
	End If
	Dim oNTCGColl As ClientGraphicsCollection = oPDoc.NonTransactingClientGraphicsCollection
	For Each oCG As Inventor.ClientGraphics In oNTCGColl
		Try
			oCG.Delete()
		Catch ex As Exception
			Logger.Error(ex.ToString())
		End Try
	Next
	Dim oCGColl As Inventor.ClientGraphicsCollection = oPDoc.ComponentDefinition.ClientGraphicsCollection
	If oCGColl IsNot Nothing AndAlso oCGColl.Count > 0 Then
		For Each oCG As Inventor.ClientGraphics In oCGColl
			Try
				oCG.Delete()
			Catch ex As Exception
				Logger.Error(ex.ToString())
			End Try
		Next
	End If
	oInvApp.ActiveView.Update()
End Sub