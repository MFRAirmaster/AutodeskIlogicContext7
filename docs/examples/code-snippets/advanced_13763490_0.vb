' Title: How to &quot;Mark&quot; the part number and revision on a part using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-quot-mark-quot-the-part-number-and-revision-on-a-part/td-p/13763490#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:38:56.931530

Dim markText = iProperties.Value("Project", "Part Number") & " - " & iProperties.Value("Project", "Revision Number")