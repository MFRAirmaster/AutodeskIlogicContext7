' Title: Rule with multiple options
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/rule-with-multiple-options/td-p/13783705
' Category: api
' Scraped: 2025-10-07T14:07:55.595605

Active_Condition = (DoorType = "3F" And (DoorWidthOptions = "2-6" Or DoorWidthOptions = "2-8" Or DoorWidthOptions = "2-10" Or DoorWidthOptions = "3-0"))

Component.IsActive("1K Skin: Back") = Active_Condition
Component.IsActive("1K Skin: Front") = Active_Condition
Component.IsActive("2K Skin: Back") = Not Active_Condition
Component.IsActive("2K Skin: Front") = Not Active_Condition
Component.IsActive("3F Skin: Back") = Not Active_Condition
Component.IsActive("3F Skin: Front") = Not Active_Condition
Component.IsActive("3F (3-6) Skin: Back") = Not Active_Condition
Component.IsActive("3F (3-6) Skin: Front") = Not Active_Condition
Component.IsActive("3S Skin: Back") = Not Active_Condition
Component.IsActive("3S Skin: Front") = Not Active_Condition
Component.IsActive("50 Skin: Back") = Not Active_Condition
Component.IsActive("50 Skin: Front") = Not Active_Condition