' Title: Utilizing MakePath in SelectSet.Select
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/utilizing-makepath-in-selectset-select/td-p/13835725
' Category: advanced
' Scraped: 2025-10-09T09:05:59.777877

Dim odoc As AssemblyDocument = ThisDoc.Document
Dim occ_proxy As ComponentOccurrenceProxy = odoc.ComponentDefinition.Occurrences.Item(2).SubOccurrences.Item(1)
odoc.SelectSet.Select(occ_proxy)
Logger.Info(odoc.SelectSet(1).name)