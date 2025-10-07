' Title: Suppress/make invisible weldment features with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/suppress-make-invisible-weldment-features-with-ilogic/td-p/13826974#messageview_0
' Category: advanced
' Scraped: 2025-10-07T12:57:36.927984

Dim oDoc As AssemblyDocument = ThisDoc.Document
Dim oWeldDef As WeldmentComponentDefinition = oDoc.ComponentDefinition

oWeldDef.Welds.Item("Weld 1").Visible = True
oWeldDef.Welds.Item("Weld 2").Visible = False