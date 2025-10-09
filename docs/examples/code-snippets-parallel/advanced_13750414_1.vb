' Title: setting Base Quantity to m^2 with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/setting-base-quantity-to-m-2-with-ilogic/td-p/13750414
' Category: advanced
' Scraped: 2025-10-09T08:55:03.231616

Parameter("VOL") = iProperties.Volume
Parameter("AREA_INS") = iProperties.Volume/50.0 mm

iProperties.Value("Custom","MATERIAL LIST DESCRIPTION") = "50mm BATTS"