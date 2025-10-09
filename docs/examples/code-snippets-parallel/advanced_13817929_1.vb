' Title: Editing feature definitions after creation?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/editing-feature-definitions-after-creation/td-p/13817929
' Category: advanced
' Scraped: 2025-10-09T08:54:34.390233

Sub main()
	
	Dim App As Inventor.Application = ThisApplication
	Dim oDoc As PartDocument = App.ActiveDocument
	Dim CompDef As PartComponentDefinition = oDoc.ComponentDefinition
	Dim oTO As TransientObjects = App.TransientObjects
	
	Dim oParams As Inventor.UserParameters = CompDef.Parameters.UserParameters
	
	Dim oShell As ShellFeature = CompDef.Features.ShellFeatures.Item("F2 Shell feature")
	Dim oShellDef As ShellDefinition = oShell.Definition
	
	Dim oCombine As CombineFeature = CompDef.Features.CombineFeatures.Item("F2 combine")

	Dim FaceCol As FaceCollection = oTO.CreateFaceCollection
	
	For Each iFace As Face In oCombine.Faces
		If iFace.Evaluator.Area < 1 Then 
			FaceCol.Add(iFace)
		End If 
	Next
	
	oShell.SetEndOfPart(True)
	
	For Each iFace As Face In FaceCol
		oShellDef.InputFaces.Add(iFace)
	Next
	
	CompDef.SetEndOfPartToTopOrBottom(False)


	

End Sub