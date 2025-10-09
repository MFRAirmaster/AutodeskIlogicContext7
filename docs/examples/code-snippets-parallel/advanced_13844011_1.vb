' Title: DXF method with INI not working, sOut version works
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/dxf-method-with-ini-not-working-sout-version-works/td-p/13844011
' Category: advanced
' Scraped: 2025-10-09T09:05:10.645616

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
	
	MsgBox(	oDataMedium.FileName )

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