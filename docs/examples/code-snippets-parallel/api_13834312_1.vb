' Title: Check Vault Status
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/check-vault-status/td-p/13834312
' Category: api
' Scraped: 2025-10-07T14:09:22.791235

Imports Autodesk.DataManagement.Client.Framework.Vault
Imports Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports AWS = Autodesk.Connectivity.WebServices
Imports AWST = Autodesk.Connectivity.WebServicesTools
Imports VB = Connectivity.Application.VaultBase
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency'.Properties
AddReference "Autodesk.DataManagement.Client.Framework.Vault.dll"
AddReference "Autodesk.DataManagement.Client.Framework.dll"
AddReference "Connectivity.Application.VaultBase.dll"
AddReference "Autodesk.Connectivity.WebServices.dll"
Sub Main
	Dim sLocalFile As String = ThisDoc.PathAndFileName(True)
	Logger.Info("sLocalFile = " & sLocalFile)
	Dim bIsLatestVersion As Boolean = LocalFileIsLatestVersion(sLocalFile)
	MsgBox("Specified File Is Latest Version = " & bIsLatestVersion.ToString(), , "")
End Sub

Function LocalFileIsLatestVersion(fullFileName As String) As Boolean
	'validate the input
	If String.IsNullOrWhiteSpace(fullFileName) Then
		Logger.Debug("Empty value passed into the 'LocalFileIsLatestVersion' method!")
		Return False
	End If
	If Not System.IO.File.Exists(fullFileName) Then
		Logger.Debug("File specified as input into the 'LocalFileIsLatestVersion' method does not exist!")
		Return False
	End If

	'get Vault connection
	Dim oConn As VDF.Vault.Currency.Connections.Connection
	oConn = VB.ConnectionManager.Instance.Connection

	'if no connection then exit routine
	If oConn Is Nothing Then
		Logger.Debug("Could not get Vault Connection!")
		MessageBox.Show("Not Logged In to Vault! - Login first and repeat executing this rule.")
		Exit Function
	End If

	'convert the 'local' path to a 'Vault' path
	Dim sVaultPath As String = iLogicVault.ConvertLocalPathToVaultPath(fullFileName)
	Logger.Info("sVaultPath = " & sVaultPath)

	'get the 'Folder' object
	Dim oFolder As AWS.Folder = oConn.WebServiceManager.DocumentService.GetFolderByPath(sVaultPath)
	If oFolder Is Nothing Then
		Logger.Warn("Did NOT get the Vault Folder that this file was in.")
	Else
		Logger.Info("Got the Vault Folder that this file was in.")
	End If
	
	'get a dictionary of all property definitions
	'if we provide an 'EmptyCategory' instead of 'Nothing', we get no entries
	Dim oPropDefsDict As VDF.Vault.Currency.Properties.PropertyDefinitionDictionary
	oPropDefsDict = oConn.PropertyManager.GetPropertyDefinitions( _
	VDF.Vault.Currency.Entities.EntityClassIds.Files, _
	Nothing, _
	VDF.Vault.Currency.Properties.PropertyDefinitionFilter.IncludeAll)
	
	If (oPropDefsDict Is Nothing) OrElse (oPropDefsDict.Count = 0) Then
		Logger.Warn("PropertyDefinitionDictionary was either Nothing or Empty!")
	Else
		Logger.Info("Got the PropertyDefinitionDictionary, and it had " & oPropDefsDict.Count.ToString() & " entries.")
	End If

	'get the VaultStatus PropertyDefinition
	Dim oStatusPropDef As VDF.Vault.Currency.Properties.PropertyDefinition
	oStatusPropDef = oPropDefsDict(VDF.Vault.Currency.Properties.PropertyDefinitionIds.Client.VaultStatus)

	If (oStatusPropDef Is Nothing) Then
		Logger.Warn("PropertyDefinition was Nothing!")
	Else
		Logger.Info("Got the PropertyDefinition OK.")
	End If

	'get all child File objects in that folder, except the 'hidden' ones
	Dim oChildFiles As AWS.File() = oConn.WebServiceManager.DocumentService.GetLatestFilesByFolderId(oFolder.Id, False)
	'if no child files found, then exit routine
	If (oChildFiles Is Nothing) OrElse (oChildFiles.Length = 0) Then
		Logger.Warn("No child files in specified folder!")
		Return False
	Else
		Logger.Info("Found " & oChildFiles.Length.ToString() & " child files in specified folder!")
	End If

	'start iterating through each file in the folder
	For Each oFile As AWS.File In oChildFiles
		'only process one specific file, by its file name
		If Not oFile.Name = System.IO.Path.GetFileNameWithoutExtension(fullFileName) Then Continue For
		Logger.Info("Found the matching Vault file.")
		
		'get the FileIteration object of this File object
		Dim oFileIt As VDF.Vault.Currency.Entities.FileIteration
		oFileIt = New VDF.Vault.Currency.Entities.FileIteration(oConn, oFile)
		If oFileIt IsNot Nothing Then
			Logger.Info("Got the FileIteration for this file.")
		Else
			Logger.Warn("Did NOT get the FileIteration for this file.")
			Continue For
		End If
		
		'Dim oPropExtProv As VDF.Vault.Interfaces.IPropertyExtensionProvider
		'oPropExtProv = oConn.PropertyManager.
		
		'Dim oPropValueSettings As New VDF.Vault.Currency.Properties.PropertyValueSettings()
		'oPropValueSettings.AddPropertyExtensionProvider(oPropExtProv)
				
		'read value of VaultStatus Property of specified File
		Dim oStatus As VDF.Vault.Currency.Properties.EntityStatusImageInfo
		oStatus = oConn.PropertyManager.GetPropertyValue(oFileIt, oStatusPropDef, Nothing)
		
		'check value, and respond accordingly
		If oStatus.Status.VersionState = VDF.Vault.Currency.Properties.EntityStatus.VersionStateEnum.MatchesLatestVaultVersion Then
			Logger.Info("Following File Version Matches Latest Vault Version:" _
			& vbCrLf & oFile.Name)
			Return True
		Else
			Logger.Info("Following File Version Does Not Match Latest Vault Version:" _
			& vbCrLf & oFile.Name)
		End If
	Next
	Return False
End Function