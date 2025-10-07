' Title: Sample to Modify AssetValue
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/sample-to-modify-assetvalue/td-p/13811384#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:35:21.185200

Sub Main()
    Dim oDoc As PartDocument
    oDoc = ThisApplication.Documents.Add(kPartDocumentObject)
        
    ' Create a custom aslib which is read-write
    Dim oLib As AssetLibrary
    oLib = ThisApplication.AssetLibraries.Add("C:\Temp\TempLib.adsklib")
        
    ' Copy a material asfrom  Autodesk Material Library to custom lib
    Dim oLibAsset As Asset
    oLibAsset= ThisApplication.AssetLibraries.Item("AD121259-C03E-4A1D-92D8-59A22B4807AD").MaterialAssets("Material-001") ' Porcelain, Ivory
    
    oLibAsset.CopyTo(oLib)
    oLibAsset= oLib.MaterialAssets(1)
    
    Dim oLocalAsset As Asset
    oLocalAsset= oLibAsset.CopyTo(oDoc,False)
    
    Dim material As MaterialAsset
    material = oLocalAsset
    
    ' update the density value 
	Dim dDensity As Double
	dDensity = material.PhysicalPropertiesAsset.Item("structural_Density").Value
	Logger.Info("Initial density value: " & dDensity)
    material.PhysicalPropertiesAsset.Item("structural_Density").Value = material.PhysicalPropertiesAsset.Item("structural_Density").Value+ 550#
    Logger.Info("Updataed density value: " & material.PhysicalPropertiesAsset.Item("structural_Density").Value)
    
    ' Update the lib asset
    material.CopyTo (oLib, True)
    Logger.Info( "Lib density: " & oLib.MaterialAssets(1).PhysicalPropertiesAsset.Item("structural_Density").Value)
End Sub