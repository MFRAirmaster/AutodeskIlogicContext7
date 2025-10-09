' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:09:48.375498

Dim oFinishDef As FinishDefinition = oFinishFeatures.CreateFinishDefinition(oFaceCol, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppearance)