' Title: Ilogic rule created by someone who left the company not working. Help needed
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-created-by-someone-who-left-the-company-not-working/td-p/13816478
' Category: advanced
' Scraped: 2025-10-07T13:21:22.366621

'Activate current document
Dim InvApp As Inventor.Application = ThisApplication
Dim oDoc As Document = ThisDoc.Document
Dim oDocType As Inventor.DocumentTypeEnum = oDoc.DocumentType

'Exit if current document is a not and assembly
If Not oDocType = DocumentTypeEnum.kAssemblyDocumentObject Then	
	Exit Sub
End If

Dim oDef As AssemblyComponentDefinition = oDoc.componentdefinition
Dim AppearanceLibraryName As String = "Rootwave Finish Library"
Dim AppLib As AssetLibrary = Nothing

Try
	AppLib = InvApp.AssetLibraries.Item(AppearanceLibraryName)
Catch
	MsgBox("Couldn't find a appearance library called '" & AppearanceLibraryName,, "Error")
	Exit Sub
End Try

Dim AppColours As AssetsEnumerator = AppLib.AppearanceAssets
Dim ColourArray As New ArrayList

'hide error
InvApp.SilentOperation = True

' Write all the appearance names to an arraylist
For Each A As Asset In AppColours
    ColourArray.Add(A.DisplayName)
Next

 'Sort the array alphabetically
ColourArray.Sort()

' Create 3 parameters with the list of appearances in......
 'Parameter 1
Try
MultiValue.List("RootwaveFinish") = ColourArray
Catch
 'Parameter doesn't exist - create it
Dim oParam As Parameter = oDef.Parameters.UserParameters.AddByValue("RootwaveFinish", ColourArray(0), UnitsTypeEnum.kTextUnits)
MultiValue.List("RootwaveFinish") = ColourArray
End Try
InvApp.SilentOperation = False