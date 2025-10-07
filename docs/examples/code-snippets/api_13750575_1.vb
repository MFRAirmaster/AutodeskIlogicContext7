' Title: How to determine a simplified part via iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-determine-a-simplified-part-via-ilogic/td-p/13750575
' Category: api
' Scraped: 2025-10-07T13:02:49.338885

If oDef.ReferenceComponents.ShrinkwrapComponents.Count > 0 Then
    isSimplified = True
End If