' Title: iProperty PropertySets Cheat Sheet, Autodesk page is down
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/iproperty-propertysets-cheat-sheet-autodesk-page-is-down/td-p/13798666#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:56:50.059101

Dim invDoc As Document
invDoc = ThisApplication.ActiveDocument

' Get the property sets
Dim DesignTrackPropSet As PropertySet = invDoc.PropertySets.Item("Design Tracking Properties")
Dim AppSummaryPropSet As PropertySet = invDoc.PropertySets.Item("Inventor Summary Information")
Dim DocSummaryPropSet As PropertySet = invDoc.PropertySets.Item("Inventor Document Summary Information")
Dim CustomPropSet As PropertySet = invDoc.PropertySets.Item("Inventor User Defined Properties")
'get a property value using the enum/prop ID
Dim Author = AppSummaryPropSet.ItemByPropId(kAuthorSummaryInformation).Value

MsgBox(Author)