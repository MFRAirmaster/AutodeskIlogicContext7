' Title: Suppress/make invisible weldment features with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/suppress-make-invisible-weldment-features-with-ilogic/td-p/13826974
' Category: advanced
' Scraped: 2025-10-07T14:27:52.509548

Dim oDoc As AssemblyDocument = ThisDoc.Document
Dim oWeldDef As WeldmentComponentDefinition = oDoc.ComponentDefinition

oWeldDef.Welds.Item("Weld 1").Visible = True
oWeldDef.Welds.Item("Weld 2").Visible = False