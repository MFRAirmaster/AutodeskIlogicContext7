' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:54:57.689298

Dim oFinishDef As FinishDefinition = oFinishFeatures.CreateFinishDefinition(oFaceCol, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppearance)