' Title: Cut a solid by RevolveFeature
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/cut-a-solid-by-revolvefeature/td-p/13837083
' Category: advanced
' Scraped: 2025-10-09T08:59:50.655940

Dim objects As ObjectCollection = oInvApp.TransientObjects.CreateObjectCollection
Dim revolveFeatRemove2 As RevolveFeature
revolveFeatRemove2 = DefinizioneParte.Features.RevolveFeatures.AddByAngle(oProfile01, oWorkAx, oElbowAngle, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, PartFeatureOperationEnum.kCutOperation)

objects.Clear

For Each body As SurfaceBody In DefinizioneParte.SurfaceBodies
If body IsNot revolveFeatRemove2.SurfaceBodies(1) Then  objects.Add(body)
Next
revolveFeatRemove2.SetAffectedBodies(objects)