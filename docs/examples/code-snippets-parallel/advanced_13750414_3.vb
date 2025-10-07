' Title: setting Base Quantity to m^2 with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/setting-base-quantity-to-m-2-with-ilogic/td-p/13750414
' Category: advanced
' Scraped: 2025-10-07T14:00:46.899939

' Check if "VOL" exists
Dim found As Boolean = False
For Each param As Parameter In ThisDoc.Document.ComponentDefinition.Parameters.UserParameters
    If param.Name = "VOL" Then
        found = True
        Exit For
    End If
Next

' Create VOL if missing, using unit string instead of enum
If Not found Then
    ThisDoc.Document.ComponentDefinition.Parameters.UserParameters.AddByValue("VOL", 1, "m^3")
End If

' Assign volume to VOL
VOL = iProperties.Volume

' Check if "AREA_INS" exists
Dim found2 As Boolean = False
For Each param As Parameter In ThisDoc.Document.ComponentDefinition.Parameters.UserParameters
    If param.Name = "AREA_INS" Then
        found2 = True
        Exit For
    End If
Next

' Create AREA_INS if missing, using unit string instead of enum
If Not found2 Then
    ThisDoc.Document.ComponentDefinition.Parameters.UserParameters.AddByValue("AREA_INS", 1, "m^2")
End If

' Assign area to AREA_INS
AREA_INS = iProperties.Volume/50 mm


' Assign New Parameters
Parameter("VOL") = iProperties.Volume
Parameter("AREA_INS") = iProperties.Volume/50 mm

iProperties.Value("Custom","MATERIAL LIST DESCRIPTION") = "50mm BATTS"