' Title: setting Base Quantity to m^2 with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/setting-base-quantity-to-m-2-with-ilogic/td-p/13750414
' Category: advanced
' Scraped: 2025-10-09T08:55:03.231616

Dim oPDoc As PartDocument = TryCast(ThisDoc.Document, PartDocument)
If oPDoc Is Nothing Then Exit Sub
Dim oPDef As PartComponentDefinition = oPDoc.ComponentDefinition
Dim dVOL As Double = oPDef.MassProperties.Volume
If dVOL = 0 Then Exit Sub
Dim oUsParams As UserParameters = oPDef.Parameters.UserParameters

Try
	oUsParams("VOL").Value = dVOL
Catch
	Dim oUsParam As UserParameter = oUsParams.AddByValue("VOL", dVOL, "m m m")
End Try

Try
	oUsParams("AREA_INS").Value = dVOL / 5
Catch
	Dim oUsParam As UserParameter = oUsParams.AddByValue("AREA_INS", dVOL / 5, "m m")
	oUsParam.ExposedAsProperty = True
	oUsParam.Units = "m"
	Call oPDef.BOMQuantity.SetBaseQuantity(BOMQuantityTypeEnum.kParameterBOMQuantity, oUsParam)
	oUsParam.Units = "m m"
End Try
Dim oCustom As PropertySet = oPDoc.PropertySets("Inventor User Defined Properties")
iProperties.Value("Custom","MATERIAL LIST DESCRIPTION") = "50mm BATTS"