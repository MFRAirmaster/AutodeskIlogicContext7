' Title: View &quot;Dim&quot; Variable Values
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/view-quot-dim-quot-variable-values/td-p/13332453
' Category: basic
' Scraped: 2025-10-07T13:55:26.144749

Dim Val1 As Double = 1.5
Dim Val2 As Double = 3
Dim Val3 As Double = 5.5
Dim values As New List(Of Double) From {Val1, Val2, Val3 }
Dim sum As Double = values.Sum
Logger.Info("Sum = " & sum.ToString())