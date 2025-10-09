' Title: Is there a faster method than DataIO.WriteDataToFile for exporting DXF's?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/is-there-a-faster-method-than-dataio-writedatatofile-for/td-p/13746690
' Category: advanced
' Scraped: 2025-10-09T09:03:47.327056

Public Class ThisRule
	
Private vaultAddinID As String = "{48B682BC-42E6-4953-84C5-3D253B52E77B}"
Private vaultAddin As ApplicationAddIn

	Sub Main()
	
	vaultAddin = ThisApplication.ApplicationAddIns.ItemById(vaultAddinID)
	vaultAddin.Deactivate()
	
	On Error Resume Next
	
	'check that the active document is an assembly file
	If ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
	MessageBox.Show ("Please run this rule from an assembly containing sheet metal parts.", "Autodesk Inventor", MessageBoxButtons.OK, MessageBoxIcon.Error)
	vaultAddin.Activate()
	Exit Sub
	End If

	Dim oBOM As BOM
	Dim oDoc As Document
	Dim oAsmDoc As AssemblyDocument
	Dim oFileDir As String
	Dim PQty As String
	Dim oPartsOnlyBOMView As BOMView
	Dim oBomRow As BOMRow
	Dim oFace As Face
	Dim oFace1 As String
	Dim oEdge As Edge
	Dim oEvaluator As CurveEvaluator
	Dim minparam As Double
	Dim maxparam As Double
	Dim length As Double
	Dim Laser As Double
	Dim LaserInchs As Double
	Dim oLaser As String
	Dim oSelectSet As SelectSet
	Dim Area As Double
	Dim oArea As Double
	Dim oSumArea As Double
	Dim counter As Integer
	Dim Holes As Integer
	Dim oDTProps As PropertySet
	Dim oWidth As String
	Dim oLength As String
	Dim MSdata(10000, 10)
	Dim sFName As String
	Dim failcounter As Integer
	Dim newcounter As Integer
	Dim sOut As String
	Dim FailedPart(10000)
	Dim i As Integer
	Dim MsgString As String
	Dim oDataIO As DataIO
	Dim oMat As String
	Dim path As String
	Dim UserName As String
	Dim oSheetMetalCompDef As SheetMetalComponentDefinition
	Dim fp As FlatPattern
	Dim CivilianHour As String
	Dim AMorPM As String
	Dim oDocType As AssemblyDocument
	Dim asmTemplate As String
	Dim TempAsm As AssemblyDocument
	Dim position As Matrix
	Dim occ As ComponentOccurrence
	Dim DocPath As String
	Dim AName As String
	Dim oBends As FlatBendResults
	Dim Bends As Integer	
	Dim Formed As String
	Dim AnythingFormed As String

	'creates DXF_CABINET folder for outputting files if it doesn't already exist on the users C drive
	path = "C:\DXF_CABINET\"
	If Len(Dir(path, vbDirectory)) = 0 Then
	    MkDir (path)
	    Else 'do nothing
	End If

	Dim Date1 As String
	Dim Hour As String
	Dim Time As DateTime = DateTime.Now
	Dim Format As String = "MM-dd-yyyy"
	Date1 = Time.ToString(Format)
	Hour = Time.ToString("HH")
	Dim MinAndSec As String = Time.ToString("mm.ss")
	If Hour = "13" Then
	    CivilianHour = "1"
	    AMorPM = "PM"
	ElseIf Hour = "14" Then
		CivilianHour = "2"
		AMorPM = "PM"
	ElseIf Hour = "15" Then
		CivilianHour = "3"
		AMorPM = "PM"	
	ElseIf Hour = "16" Then
		CivilianHour = "4"
		AMorPM = "PM"	
	ElseIf Hour = "17" Then
		CivilianHour = "5"
		AMorPM = "PM"	
	ElseIf Hour = "18" Then
		CivilianHour = "6"
		AMorPM = "PM"	
	ElseIf Hour = "19" Then
		CivilianHour = "7"
		AMorPM = "PM"	
	ElseIf Hour = "20" Then
		CivilianHour = "8"
		AMorPM = "PM"	
	ElseIf Hour = "21" Then
		CivilianHour = "9"
		AMorPM = "PM"	
	ElseIf Hour = "22" Then
		CivilianHour = "10"
		AMorPM = "PM"	
	ElseIf Hour = "23" Then
		CivilianHour = "11"
		AMorPM = "PM"	
	Else
	    CivilianHour = Hour
	    AMorPM = "AM"
	End If
	
	'establish top level part number for the folder/file name
	FolderNumber = ThisApplication.ActiveDocument.PropertySets(3).Item("Part Number").Value


	path = "C:\DXF_CABINET\" & FolderNumber & "  " & Date1 & " at " & CivilianHour & "." & MinAndSec & " " & AMorPM
	MkDir (path)
	oFileDir = path & "\"

	oDoc = ThisApplication.ActiveDocument
	
	oAsmDoc = ThisApplication.ActiveDocument
		oDocType = ThisApplication.ActiveDocument
		
		oAsmNameProp = oAsmDoc.PropertySets(3).Item("Part Number")
		AName = oAsmNameProp.Value
		
		'check if its an iassembly and create a temporary assembly with the active member if so
		If oAsmDoc.ComponentDefinition.IsiAssemblyFactory = True Then
		    DocPath = oAsmDoc.FullFileName
		    asmTemplate = ThisApplication.FileManager.GetTemplateFile(kAssemblyDocumentObject) 'Your assembly template file here
		    TempAsm = ThisApplication.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, asmTemplate, False)
		    position = ThisApplication.TransientGeometry.CreateMatrix()
		    occ = TempAsm.ComponentDefinition.Occurrences.AddiAssemblyMember(DocPath, position)
		    oAsmDoc = TempAsm
		End If
	
		oBOM = oAsmDoc.ComponentDefinition.BOM
		
		'Make sure that the parts only view is enabled.
		oBOM.PartsOnlyViewEnabled = True
		
		'Change Bom structure to Normal for all that are inseparable (doing this because weldments have an inseparable bom structure and we need to look at the bom)
		Dim ARD As Document
		For Each ARD In oAsmDoc.AllReferencedDocuments
		    If ARD.DocumentType = kAssemblyDocumentObject Or kPartDocumentObject Then
		        If ARD.ComponentDefinition.BOMStructure = Inventor.BOMStructureEnum.kInseparableBOMStructure Then
		            ARD.ComponentDefinition.BOMStructure = Inventor.BOMStructureEnum.kNormalBOMStructure
		        End If
		    End If
		Next ARD
		    
		' Set a reference to the "Parts Only" BOMView
		oPartsOnlyBOMView = oBOM.BOMViews.Item("Parts Only")
		
		For Each oBomRow In oPartsOnlyBOMView.BOMRows
		
		    oDoc = oBomRow.ComponentDefinitions(1).Document
		      
		    If oDoc.ComponentDefinition.Type = kSheetMetalComponentDefinitionObject Then
	        
	            oSheetMetalCompDef = oDoc.ComponentDefinition

	            If oSheetMetalCompDef.HasFlatPattern = False Then
	                oSheetMetalCompDef.Unfold
	                oSheetMetalCompDef.FlatPattern.ExitEdit
	            End If
	           
			    If oSheetMetalCompDef.HasFlatPattern = False Then  'after a fail checks if the part is a member of a factory which would be why it couldnt unbend
						If oDoc.ComponentDefinition.IsiPartMember Then
							
							Dim oPDoc As PartDocument
							oPDoc = oDoc
							Dim oPDef As PartComponentDefinition = oPDoc.ComponentDefinition
							Dim oPFactory As iPartFactory = Nothing
							
							oPFactory = oPDef.iPartMember.ParentFactory
							Dim oFactoryPDoc As PartDocument = oPFactory.Parent
							
							oSheetMetalCompDef = oFactoryPDoc.ComponentDefinition
							oSheetMetalCompDef.Unfold
		                	oSheetMetalCompDef.FlatPattern.ExitEdit
							
							Dim oRow As iPartTableRow
							For Each oRow In oFactoryPDoc.ComponentDefinition.iPartFactory.TableRows
							oFactoryPDoc.ComponentDefinition.iPartFactory.CreateMember(oRow)
							Next
							
							oDoc = oBomRow.ComponentDefinitions(1).Document

						End If
		        End If
			   
	            If oSheetMetalCompDef.HasFlatPattern = False Then  'records which parts failed to be displayed upon completion of ilogic if there are failed parts
	                oPartNameProp = oDoc.PropertySets(3).Item("Part Number")
	                sFName = oPartNameProp.Value
	                    If sFName = "" Then
	                        sFName = Replace(oDoc.DisplayName, ".ipt", "")
	                    End If
	               FailedPart(failcounter) = sFName
	               failcounter = failcounter + 1
	            Else
	            
	                fp = oSheetMetalCompDef.FlatPattern
	            
	                ' Get the DataIO object.
	                oDataIO = oDoc.ComponentDefinition.DataIO
	                
	                ' Set FileName
	                oPartNameProp = oDoc.PropertySets(3).Item("Part Number")
	                sFName = oPartNameProp.Value
	                    If sFName = "" Then
	                        sFName = Replace(oDoc.DisplayName, ".ipt", "")
	                    End If
	                
	                ' Build the string that defines the format of the DXF file.
	                sOut = "FLAT PATTERN DXF?AcadVersion=R12&SimplifySplines=true&FeatureProfilesUpLayerCOLOR=255;255;0&InvisibleLayers=IV_BEND;IV_BEND_DOWN;IV_TANGENT;IV_ARC_CENTERS;IV_FEATURE_PROFILES_DOWN;IV_TOOL_CENTER;IV_TOOL_CENTER_DOWN;IV_ROLL_TANGENT;IV_ALTREP_BACK;IV_ROLL;"
	                oMat = oDoc.PropertySets(3).Item("Material").Value
	                If InStr(oMat, "1/4 HSLA") > 0 Then
	                    oMat = Replace(oMat, "1/4", "1_4")
	                End If

'///////***************************************************************************//////////
'///////***************************************************************************//////////

 					oDataIO.WriteDataToFile( sOut, oFileDir & sFName & " - " & oMat & ".dxf")

'///////***************************************************************************//////////
'///////***************************************************************************//////////

	            End If
			End If
		Next
				
				If failcounter > 1 Then     'displays list of failed parts if there are any
		        While i < failcounter
		            MsgString = MsgString & FailedPart(i) & vbCr
		            i = i + 1
		        End While
			Dim Title As Object = "Attention"	
		    MessageBox.Show (MsgString,"Couldn't Unbend:  ", MessageBoxButtons.OK, MessageBoxIcon.Error)
		End If
		
		If oDocType.ComponentDefinition.IsiAssemblyFactory = True Then
		TempAsm.Close
		End If


	Shell("C:\WINDOWS\explorer.exe """ & oFileDir & "", vbNormalFocus) 'opens folder for user at finish

	vaultAddin.Activate()
	
	End Sub
End Class