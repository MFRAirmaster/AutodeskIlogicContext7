' Title: Trailing Zeros
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/trailing-zeros/td-p/13803427#messageview_0
' Category: api
' Scraped: 2025-10-09T09:01:47.917202

Dim Height As Double = tank_id
Dim Width As Double = tank_height

Dim HeightString As String = CStr(Height)
Dim WidthString As String = CStr(Width)

Dim HeightParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_id")
Dim WidthParam As Inventor.Parameter = ThisApplication.ActiveDocument.componentDefinition.Parameters.Item("tank_height")

HeightParam.Precision = 3 'I wish this worked to control the decimals displayed from either the expression or value. 
WidthParam.Precision = 3

Dim HeightSubString As String() = HeightParam.Expression.Split(" "c) 
Dim WidthSubString As String() = WidthParam.Expression.Split(" "c)


Logger.Info("Height blue value: " & tank_id & " Height double: " & Height & " Height String: " & HeightString & " Height param expression: " & HeightParam.Expression & " Height param value direct: " & HeightParam.Value / 2.54 & " Height substring: " & HeightSubString(0))
Logger.Info("Width blue value: "  & tank_height & " Widdth double: " & Width & "  Width String: " & WidthString & " Width param expression: " & WidthParam.Expression & " Width param value direct: " & HeightParam.Value / 2.54 & " Width substring: " & WidthSubString(0))