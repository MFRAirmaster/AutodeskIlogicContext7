' Title: Logic to match Inventor file names to the search engine names
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/logic-to-match-inventor-file-names-to-the-search-engine-names/td-p/13824844
' Category: advanced
' Scraped: 2025-10-07T13:32:49.500518

Dim ass As AssemblyDocument = ThisDoc.Document
For Each occ As ComponentOccurrence In ass.ComponentDefinition.Occurrences
	If occ.Suppressed Then Continue For
	If TypeOf occ.Definition Is VirtualComponentDefinition Then Continue For
	Dim occ_name As String = ""
	occ_name = Split(occ.Name, ":")(0)
	Dim full_filename As String = ""
	full_filename=occ.Definition.Document.fullfilename
	Dim path As String = ""
	path=Microsoft.VisualBasic.left(full_filename, InStrRev(full_filename, "\", -1))
    Dim file_name As String = ""
	file_name=Microsoft.VisualBasic.Right(full_filename, Len(full_filename) -InStrRev(full_filename, "\", -1))
	Dim ext As String = ""
	ext = Microsoft.VisualBasic.Right(full_filename, 4)
	file_name = Split(file_name, ".")(0)
		
	If file_name.ToLower <> occ_name.ToLower Then
	
	Dim doc As Document = occ.Definition.Document
	Logger.Info(path & "  " & file_name & "  " & occ_name)
	If Dir(path & occ_name & ext) = "" Then
		Try
		doc.SaveAs(path & occ_name & ext, True)
		Catch ex As Exception
		End Try
	End If
		Try
		occ.Replace(path & occ_name & ext,True)
		Catch ex As Exception
		End Try
	
	
	End If
Next