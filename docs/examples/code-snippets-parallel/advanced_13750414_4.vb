' Title: setting Base Quantity to m^2 with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/setting-base-quantity-to-m-2-with-ilogic/td-p/13750414
' Category: advanced
' Scraped: 2025-10-07T14:00:46.899939

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
	oUsParam.ExposedAsProperty = True
	Call oPDef.BOMQuantity.SetBaseQuantity(BOMQuantityTypeEnum.kParameterBOMQuantity, oUsParam)
End Try

Try : oUsParams("AREA_INS").Value = dVOL / 5
Catch : oUsParams.AddByValue("AREA_INS", dVOL / 5, "m m")
End Try

Dim oCustom As PropertySet = oPDoc.PropertySets("Inventor User Defined Properties")
Try : oCustom("MATERIAL LIST DESCRIPTION").Value = "50mm BATTS"
Catch : oCustom.Add("MATERIAL LIST DESCRIPTION", "50mm BATTS") : End Try