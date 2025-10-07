' Title: Create Model States using iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-model-states-using-ilogic/td-p/10714694
' Category: advanced
' Scraped: 2025-10-07T13:06:06.974609

If ManwayType = 600 Then
Feature.IsActive("MWAY600 Centre") = True
Feature.IsActive("MWAY600 Holes") = True
Feature.IsActive("MWAY600 PCD") = True
Feature.IsActive("MWAY800 Centre") = False
Feature.IsActive("MWAY800 Holes") = False
Feature.IsActive("MWAY800 PCD") = False
Feature.IsActive("MWAY900 Centre") = False
Feature.IsActive("MWAY900 Holes") = False
Feature.IsActive("MWAY800 PCD") = False

Else If ManwayType = 800 Then
Feature.IsActive("MWAY800 Centre") = True
Feature.IsActive("MWAY800 Holes") = True
Feature.IsActive("MWAY800 PCD") = True
Feature.IsActive("MWAY600 Centre") = False
Feature.IsActive("MWAY600 Holes") = False
Feature.IsActive("MWAY600 PCD") = False
Feature.IsActive("MWAY900 Centre") = False
Feature.IsActive("MWAY900 Holes") = False
Feature.IsActive("MWAY900 PCD") = False

Else If ManwayType = 900 Then
Feature.IsActive("MWAY900 Centre") = True
Feature.IsActive("MWAY900 Holes") = True
Feature.IsActive("MWAY900 PCD") = True
Feature.IsActive("MWAY800 Centre") = False
Feature.IsActive("MWAY800 Holes") = False
Feature.IsActive("MWAY800 PCD") = False
Feature.IsActive("MWAY600 Centre") = False
Feature.IsActive("MWAY600 Holes") = False
Feature.IsActive("MWAY600 PCD") = False

End If


iLogicVb.UpdateWhenDone = True