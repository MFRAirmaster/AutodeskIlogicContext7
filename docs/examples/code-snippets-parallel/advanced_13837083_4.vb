' Title: Cut a solid by RevolveFeature
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/cut-a-solid-by-revolvefeature/td-p/13837083#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:08:28.488479

Dim oInvApp As Inventor.Application=ThisApplication
Dim Parte As PartDocument = ThisDoc.Document
Dim DefinizioneParte As PartComponentDefinition=Parte.ComponentDefinition
Dim oFeatures As PartFeatures
oFeatures = DefinizioneParte.Features
Dim objects As ObjectCollection = oInvApp.TransientObjects.CreateObjectCollection

For Each body As SurfaceBody  In DefinizioneParte.SurfaceBodies
objects.Add(body)
Next

 Dim oProfile01 As Profile = oInvApp.CommandManager.Pick(SelectionFilterEnum.kSketchProfileFilter,"pick profile to revolve")
 Dim oWorkAx As WorkAxis = oInvApp.CommandManager.Pick(SelectionFilterEnum.kWorkAxisFilter,"pick axis to revolve")
 Dim oElbowAngle As Double =1

Dim revolveFeatRemove2 As RevolveFeature
revolveFeatRemove2 = DefinizioneParte.Features.RevolveFeatures.AddByAngle(oProfile01,oWorkAx,oElbowAngle,PartFeatureExtentDirectionEnum.kNegativeExtentDirection,PartFeatureOperationEnum.kCutOperation)
revolveFeatRemove2.SetAffectedBodies(objects )