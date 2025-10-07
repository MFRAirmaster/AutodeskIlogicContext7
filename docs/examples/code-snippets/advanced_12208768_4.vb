' Title: Promote a model state to become the master state with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/promote-a-model-state-to-become-the-master-state-with-ilogic/td-p/12208768
' Category: advanced
' Scraped: 2025-10-07T13:05:48.627090

AddReference  "Microsoft.Office.Interop.Excel.dll"
Imports Microsoft.Office.Interop
Imports oXL = Microsoft.Office.Interop.Excel
Imports System.Windows.Forms
Imports Inventor
Public Sub Main
	Dim oADoc As AssemblyDocument = ThisDoc.FactoryDocument
	Dim oADef As AssemblyComponentDefinition = oADoc.ComponentDefinition
	Dim oRepsMgr As RepresentationsManager = oADef.RepresentationsManager
	Dim oPosReps As PositionalRepresentations = oRepsMgr.PositionalRepresentations
	Logger.Info("PosReps Count = " & oPosReps.Count.ToString)
	Dim oActivePosRep As PositionalRepresentation = oRepsMgr.ActivePositionalRepresentation
	Logger.Info("Active PosRep Name = " & oActivePosRep.Name)
	If oPosReps.ExcelWorkSheet Is Nothing Then Return
	Dim oWS As oXL.Worksheet = TryCast(oPosReps.ExcelWorkSheet, oXL.Worksheet)
	oWS.Application.Visible = True
	'oWS.Visible = oXL.XlSheetVisibility.xlSheetVisible
	ThisApplication.UserInterfaceManager.DoEvents()
	MessageBox.Show("Review Opened Excel Worksheet for PositionalRepresentation Settings.")
	oWS = Nothing
End Sub