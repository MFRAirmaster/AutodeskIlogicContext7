' Title: Workflow to modify MaterialAsset (apply another PhysicalPropertiesAsset)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/workflow-to-modify-materialasset-apply-another/td-p/13822777
' Category: api
' Scraped: 2025-10-07T13:57:59.723326

Dim ptDoc As PartDocument = ThisDoc.Document

Logger.Info(ptDoc.MaterialAssets.Count)

Dim srcMat As Asset = ptDoc.MaterialAssets(1)
Logger.Info(srcMat.DisplayName)
Logger.Info(vbTab & srcMat.PhysicalPropertiesAsset.DisplayName)

Dim tgtMat As Asset = ptDoc.MaterialAssets(2)
Logger.Info(tgtMat.DisplayName)
Logger.Info(vbTab & tgtMat.PhysicalPropertiesAsset.DisplayName)

tgtMat.PhysicalPropertiesAsset = srcMat.PhysicalPropertiesAsset ' HRESULT: 0x80020003 (DISP_E_MEMBERNOTFOUND)