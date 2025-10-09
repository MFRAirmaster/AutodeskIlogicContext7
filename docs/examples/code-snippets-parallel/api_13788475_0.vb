' Title: how to change weld size in weld symbol by ilogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-change-weld-size-in-weld-symbol-by-ilogic/td-p/13788475
' Category: api
' Scraped: 2025-10-09T08:47:15.669205

Dim obj = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAllEntitiesFilter, "Select a weld symbol")

If (TypeOf obj IsNot DrawingWeldingSymbol) Then
    MsgBox("You did not select a welding symbol")
    Return
End If

Dim symbol As DrawingWeldingSymbol = obj
Dim def As DrawingWeldingSymbolDefinition = symbol.Definitions.Item(1)
Dim symbolDef1 = def.WeldSymbolOne
Dim symbolDef2 = def.WeldSymbolTwo

symbolDef1.Prefix = "Prefix 1"
symbolDef1.Leg1 = "Leg 1.1"
symbolDef1.Leg2 = "Leg 1.2"
symbolDef1.WeldSymbolType = WeldSymbolTypeEnum.kFilletWeldSymbolType
symbolDef1.Length = "Length 1"
symbolDef1.Pitch = "Pitch 1"

symbolDef2.Prefix = "Prefix 2"
symbolDef2.Leg1 = "Leg 2.1"
symbolDef2.Leg2 = "Leg 2.2"
symbolDef2.WeldSymbolType = WeldSymbolTypeEnum.kFilletWeldSymbolType
symbolDef2.Length = "Length 2"
symbolDef2.Pitch = "Pitch 2"