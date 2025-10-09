# DXF Export Methods for Sheet Metal Flat Patterns

This example demonstrates two methods for exporting DXF files from Inventor sheet metal parts: using an INI configuration file (troubleshooting version) and using a formatted string (working version).

## Inventor Version
Inventor 2025 (methods may work in 2024+ with adjustments).

## Description
The first method attempts to use a TranslatorAddIn with an INI file for DXF export but produces invalid files. The second method uses DataIO.WriteDataToFile with a formatted string for reliable export.

## Method 1: INI File Approach (Not Working)
This method tries to use an external INI file for DXF export configuration but fails to produce valid DXF files.

```vb
Sub Main

    DXFFolder = "C:\Temp\DXF\"
    oModelDocName = ThisDoc.Document.FullFileName
    DXFName = IO.Path.GetFileNameWithoutExtension(oModelDocName) & "-INI"

    Dim oInv As Application = ThisApplication
    Dim oDoc As Document = oInv.Documents.Open(oModelDocName, True)

    Dim oSheetDef As SheetMetalComponentDefinition = oDoc.ComponentDefinition

    If oSheetDef.HasFlatPattern = False Then
        oSheetDef.Unfold
    Else
        oSheetDef.FlatPattern.Edit
    End If

    'Set the DXF target file name
    If Not IO.Directory.Exists(DXFFolder) Then IO.Directory.CreateDirectory(DXFFolder)

    'Get the DXF translator Add-In.
    Dim DXFAddIn As TranslatorAddIn
    DXFAddIn = oInv.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}")

    'Create a DataMedium object
    Dim oDataMedium As DataMedium
    oDataMedium = oInv.TransientObjects.CreateDataMedium

    'Create a TranslationContext object
    Dim oContext As TranslationContext
    oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism

    'Get the Design Data Path
    Dim INI_Folder As String
    INI_Folder = "C:\Temp\DXF\DXF configs\"

    'Create a NameValueMap object
    Dim oOptions As NameValueMap
    oOptions = oInv.TransientObjects.CreateNameValueMap

    'Specify the INI file path for exporting
    Dim iniFilePath As String = INI_Folder & "PC_flatpatterns.ini"
    oOptions.Value("Export_Acad_IniFile") = iniFilePath

    'MsgBox(INI_Folder & "PC_flatpatterns.ini")

    'Define the file name
    oDataMedium.FileName = DXFFolder & DXFName & ".dxf"

        MsgBox(    oDataMedium.FileName )


    '    'Specify the AutoCAD version to export to (AutoCAD 2004)
    '    oOptions.Value("Export_AcadVersion") = 16

    'Publish document.
    DXFAddIn.SaveCopyAs(oDoc, oContext, oOptions, oDataMedium)



End Sub
```

## Method 2: Formatted String Approach (Working)
This method uses a formatted string to configure DXF export layers and produces valid files.

```vb
Sub Main

DXFFolder = "C:\Temp\DXF\"
oModelDocName = ThisDoc.Document.FullFileName
DXFName = IO.Path.GetFileNameWithoutExtension(oModelDocName) & "-sOUT"

    Dim oInv As Application = ThisApplication
    Dim oDoc As Document = oInv.Documents.Open(oModelDocName, False)


    Dim oSheetDef As SheetMetalComponentDefinition = oDoc.ComponentDefinition

    If oSheetDef.HasFlatPattern = False Then
        oSheetDef.Unfold
    End If

    oDataMedium = oInv.TransientObjects.CreateDataMedium

    If Not IO.Directory.Exists(DXFFolder) Then IO.Directory.CreateDirectory(DXFFolder)

    'Set the DXF target file name
    oDataMedium.FileName = DXFFolder & DXFName & ".dxf"

    MsgBox(    oDataMedium.FileName )

    Dim sOut As String = GetFormatString

    oSheetDef.DataIO.WriteDataToFile(sOut, oDataMedium.FileName)



End Sub

'set up the output string
Function GetFormatString() As String

    'https://help.autodesk.com/view/INVNTOR/2026/ENU/?guid=GUID-DataIO

    Dim FormatedString As String
    FormatedString = "FLAT PATTERN DXF?AcadVersion=2007" & _
    "&TangentLayer=IV_TANGENT " & _
    "&BendLayer=IV_BEND" & _
    "&BendDownLayer=IV_BEND_DOWN" & _
    "&ToolCenterLayer=IV_TOOL_CENTER" & _
    "&ToolCenterDownLayer=IV_TOOL_CENTER_DOWN" & _
    "&ArcCentersLayer=IV_ARC_CENTERS" & _
    "&OuterProfileLayer=IV_OUTER_PROFILE" & _
    "&InteriorProfilesLayer=IV_INTERIOR_PROFILES" & _
    "&FeatureProfilesLayer=IV_FEATURE_PROFILES" & _
    "&FeatureProfilesDownLayer=IV_FEATURE_PROFILES_DOWN" & _
    "&AltRepFrontLayer=IV_ALTREP_FRONT" & _
    "&AltRepBackLayer=IV_ALTREP_BACK" & _
    "&UnconsumedSketchesLayer=IV_UNCONSUMED_SKETCHES" & _
    "&TangentRollLinesLayer=IV_ROLL_TANGENT" & _
    "&RollLinesLayer=IV_ROLL" & _
    "&InvisibleLayers=IV_TANGENT;IV_ROLL_TANGENT" & _
    "&TangentLayerColor=0;0;0 " & _
    "&endLayerColor=255;0;0" & _
    "&BendDownLayerColor=0;255;0" & _
    "&ToolCenterLayerColor=0;0;0" & _
    "&ToolCenterDownLayerColor=0;0;0" & _
    "&ArcCentersLayerColor=0;0;0" & _
    "&OuterProfileLayerColor=0;0;0" & _
    "&InteriorProfilesLayerColor=0;0;0" & _
    "&FeatureProfilesLayerColor=0;0;0" & _
    "&FeatureProfilesDownLayerColor=0;0;0" & _
    "&AltRepFrontLayerColor=0;0;0" & _
    "&AltRepBackLayerColor=0;0;0" & _
    "&UnconsumedSketchesLayerColor=-255;-255;-255" & _
    "&TangentRollLinesLayerColor=0;0;0" & _
    "&RollLinesLayerColor=0;0;0" & _
    "&MergeProfilesIntoPolyline=True"
    Return FormatedString

End Function
```

## Usage
Use Method 2 for reliable DXF export. Ensure the sheet metal part has a flat pattern. Adjust folder paths as needed. For INI method issues, verify the INI file format matches Inventor expectations.
