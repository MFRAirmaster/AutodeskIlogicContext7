' Title: Editing feature definitions after creation?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/editing-feature-definitions-after-creation/td-p/13817929
' Category: advanced
' Scraped: 2025-10-09T08:54:34.390233

Dim oDoc As PartDocument = ThisDoc.Document
Dim oPDef As PartComponentDefinition = oDoc.ComponentDefinition
Dim oFaces As FaceCollection = ThisApplication.TransientObjects.CreateFaceCollection
Dim oBody As SurfaceBody
Dim oFace As Inventor.Face
Dim oShellFeats As ShellFeatures
Dim oShellDef As ShellDefinition
Dim oShelllFeat As ShellFeature
Dim oThickness As Double = .1

oShellFeats = oPDef.Features.ShellFeatures

' just for testing
Try
oShelllFeat = oShellFeats.Item("MyShell")
oShelllFeat.Delete 
Catch
End Try

oBody = oPDef.SurfaceBodies.Item(1)
'oFace = oBody.Faces.Item(2)
'oFaces.Add(oFace)
oShellFeats = oPDef.Features.ShellFeatures
oShellDef = oShellFeats.CreateShellDefinition(oFaces, oThickness, ShellDirectionEnum.kInsideShellDirection)
oShelllFeat = oShellFeats.Add(oShellDef)
oShelllFeat.Name = "MyShell"