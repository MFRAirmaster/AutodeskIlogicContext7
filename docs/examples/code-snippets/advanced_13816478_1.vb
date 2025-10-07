' Title: Ilogic rule created by someone who left the company not working. Help needed
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-rule-created-by-someone-who-left-the-company-not-working/td-p/13816478
' Category: advanced
' Scraped: 2025-10-07T13:21:22.366621

'Activate current document
oDocType = ThisApplication.ActiveDocumentType

'hide error
ThisApplication.SilentOperation = True

'Exit if current document is a not and assembly
If oDocType <> kAssemblyDocumentObject Then
	Return
End If

Dim oDoc As Document = ThisDoc.Document
Dim oDef As AssemblyComponentDefinition = oDoc.componentdefinition
Dim AppearanceLibraryName As String = "Rootwave Finish Library"

' Get the appearance library
Dim AppLib As AssetLibrary = Nothing
Try
	AppLib = ThisApplication.AssetLibraries.Item(AppearanceLibraryName)
Catch
	MsgBox("Couldn't find a appearance library called '" & AppearanceLibraryName,, "Error")
	Exit Sub
End Try

Dim AppColours As AssetsEnumerator = AppLib.AppearanceAssets
Dim ColourArray As New ArrayList

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
ThisApplication.SilentOperation = False