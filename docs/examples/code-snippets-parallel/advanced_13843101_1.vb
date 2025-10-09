' Title: Sheet metal cut feature affected bodies
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sheet-metal-cut-feature-affected-bodies/td-p/13843101
' Category: advanced
' Scraped: 2025-10-09T08:54:18.006574

' --- CUT FEATURE ---
	Dim oPart As PartDocument = ThisApplication.ActiveDocument
	Dim oDef As SheetMetalComponentDefinition = oPart.ComponentDefinition
	
	Dim oSMFeatures As SheetMetalFeatures = oDef.Features
	
	' --- Create cut definition ---
	Dim oCutDef As CutDefinition = oSMFeatures.CutFeatures.CreateCutDefinition(oSketch.Profiles.AddForSolid())
	
	' --- Set cut distance using the Distance variable from Excel ---
	oCutDef.SetDistanceExtent(Distance, PartFeatureExtentDirectionEnum.kNegativeExtentDirection)
	oCutDef.CutNormalToFlat = True
	
	' --- Add the cut ---
	Dim oCut As CutFeature = oSMFeatures.CutFeatures.Add(oCutDef)
	oCut.Name = "Cut_" & PointName