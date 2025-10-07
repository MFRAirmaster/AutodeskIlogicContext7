' Title: Sample to Modify AssetValue
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sample-to-modify-assetvalue/td-p/13811384#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:35:21.185200

materialasset.Item("physmat_Keywords").Value = sap_.Text
materialasset.Item("physmat_Label").Value = sap_.Text
materialasset.Item("physmat_Manufacturer").Value = Producent_.Text
materialasset.Item("physmat_Comments").Value = kategoria_.Text
materialasset.Item("physmat_Model").Value = wymiar_.Text
materialasset.Item("physmat_Type").Value = "physmat_" & assloctype.ToLower
materialasset.Item("physmat_class").Value = obr_czt.Text
materialasset.Item("physmat_Cost").Value = ilosc_.Text