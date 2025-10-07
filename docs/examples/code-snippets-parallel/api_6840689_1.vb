' Title: Run iLogic rule from Zero Document ribbon state
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/run-ilogic-rule-from-zero-document-ribbon-state/td-p/6840689
' Category: api
' Scraped: 2025-10-07T14:08:03.270053

ThisDoc.Launch("C:\TEMP\Part1.ipt")
oEditDoc = ThisApplication.ActiveEditDocument
MessageBox.Show(oEditDoc.FullFileName, "iLogic")
oEditDoc.Close