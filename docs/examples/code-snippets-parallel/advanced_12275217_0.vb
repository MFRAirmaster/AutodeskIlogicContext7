' Title: Customized Event Triggers
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/customized-event-triggers/td-p/12275217
' Category: advanced
' Scraped: 2025-10-09T08:56:03.216009

Sub Main() Dim oModelEvent As ModelingEvents  oModelEvent = ThisApplication.ModelingEvents   AddHandler oModelEvent.OnParameterChange, AddressOf oModelEvent_ParameterChangeEnd Sub Sub oModelEvent_ParameterChange(DocumentObject As _Document, Parameter As Parameter, BeforeOrAfter As EventTimingEnum, Context As NameValueMap, ByRef HandlingCode As HandlingCodeEnum)  If BeforeOrAfter = kBefore Then MessageBox.Show("Before changing parameter : " & Parameter.Value) ElseIf BeforeOrAfter = kAfter Then MessageBox.Show("After changing parameter : " & Parameter.Value) End If   End Sub