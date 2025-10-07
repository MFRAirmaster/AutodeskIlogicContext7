' Title: Apply Finish to Sub Assembly within Assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/apply-finish-to-sub-assembly-within-assembly/td-p/13753431
' Category: advanced
' Scraped: 2025-10-07T13:07:50.456065

Dim oFinishDef As FinishDefinition = oFinishFeatures.CreateFinishDefinition(oFaceCol, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppearance)