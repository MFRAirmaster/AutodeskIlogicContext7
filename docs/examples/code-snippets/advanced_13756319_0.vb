' Title: Feature Rectangular Pattern creation in assembly using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/feature-rectangular-pattern-creation-in-assembly-using-ilogic/td-p/13756319#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:23:18.884744

Dim oAsm As AssemblyDocument = ThisApplication.ActiveDocument
Dim oDef As AssemblyComponentDefinition = oAsm.ComponentDefinition

Dim oExtFeat As ExtrudeFeature = oDef.Features.Item("slot_top_cut")

Dim patt As RectangularPatternFeature

Dim oColl As ObjectCollection
oColl = ThisApplication.TransientObjects.CreateObjectCollection()
oColl.Add(oExtFeat)

Dim oAxis As WorkAxis
oAxis = oDef.WorkAxes.Item("X Axis")

Dim rectDef As RectangularPatternFeatureDefinition
rectDef = oDef.Features.RectangularPatternFeatures.CreateDefinition(oColl, oAxis, True, tube_qty_tube - 1, pitch / 10)

patt = oDef.Features.RectangularPatternFeatures.AddByDefinition(rectDef)

patt.Name = "slot_top_pattern"