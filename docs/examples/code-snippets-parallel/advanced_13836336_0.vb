' Title: How to link the material description directly from the library to the BOM in the drawing.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-link-the-material-description-directly-from-the-library/td-p/13836336#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:01:25.706198

On Error Resume Next 
Dim mat As MaterialAsset = ThisDoc.Document.activematerial
Dim value As AssetValue
			For Each value In mat
				Logger.Info(value.DisplayName & "  " & value.Name)
				iProperties.Value("Custom", "material_" & value.DisplayName)=value.value
			Next