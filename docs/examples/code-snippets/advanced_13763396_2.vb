' Title: Unable to apply finish to multiple of the same part within an assembly
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/unable-to-apply-finish-to-multiple-of-the-same-part-within-an/td-p/13763396
' Category: advanced
' Scraped: 2025-10-07T13:33:39.862350

Dim oFinishDef As FinishDefinition = oFinishFeats.CreateFinishDefinition(oFaceColl, FinishTypeEnum.kMaterialCoatingFinishType, "Powder Coat", oAppearance)