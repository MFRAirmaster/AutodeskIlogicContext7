' Title: Cut a solid by RevolveFeature
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/cut-a-solid-by-revolvefeature/td-p/13837083
' Category: advanced
' Scraped: 2025-10-07T13:34:40.078144

Hi, how can I choose the solid to cut with RevolveFeature?This is my code: Dim revolveFeatRemove As RevolveFeature = DefinizioneParte.Features.RevolveFeatures.AddByAngle(oProfile01, oWorkAx, oElbowAngle, PartFeatureExtentDirectionEnum.kNegativeExtentDirection, PartFeatureOperationEnum.kCutOperation)Thanks in advance