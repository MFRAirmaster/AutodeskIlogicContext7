' Title: Batch Plot IDWs to single PDF
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/batch-plot-idws-to-single-pdf/td-p/13437520
' Category: advanced
' Scraped: 2025-10-09T08:59:56.844603

AddReference "C:\Users\Public\global references\itextsharp.dll"
AddReference "System.IO"

Class ThisRule


	Shared _Folder_Path As String = String.Empty
	Shared _PDF_Name As String = Nothing
	Shared _GetPDF As String = String.Empty

	Sub Main()

		'set the name of the new pdf you want to create without the extension
		_PDF_Name = IO.Path.GetFileNameWithoutExtension(thisdoc.document.fullfilename) & "_Combined"

		'set the path of the folder contaiing all of the pdfs
		_Folder_Path = "C:\Temp\My PDFs"

		If Not System.IO.Directory.Exists(_Folder_Path) Then
			System.IO.Directory.CreateDirectory(_Folder_Path)
		End If

		'call the sub to created PDFs for all drawings from the current assembly
		Call PDF_ALL_Docs

		'get the PDF result
		Dim OpenPDF As String = CreatePDF(_Folder_Path)

	End Sub



	Sub PDF_ALL_Docs
		'define the active document as an assembly file
		Dim oAsmDoc As AssemblyDocument
		oAsmDoc = ThisApplication.ActiveDocument
		oAsmName = Left(oAsmDoc.DisplayName, Len(oAsmDoc.DisplayName) -4)

		'check that the active document is an assembly file
		If ThisApplication.ActiveDocument.DocumentType <> kAssemblyDocumentObject Then
			MessageBox.Show("Please run this rule from the assembly file.", "iLogic")
			Exit Sub
		End If

		'get user input
		RUsure = MessageBox.Show( _
		"This will create a PDF file for all of the asembly components that have drawings files." _
		& vbLf & "This rule expects that the drawing file shares the same name and location as the component." _
		& vbLf & " " _
		& vbLf & "Are you sure you want to create PDF Drawings for all of the assembly components?" _
		& vbLf & "This could take a while.", "iLogic  - Batch Output PDFs ", MessageBoxButtons.YesNo)

		If RUsure = vbNo Then
			Return
		Else
		End If

		'- - - - - - - - - - - - -PDF setup - - - - - - - - - - - -
		oPath = ThisDoc.Path
		PDFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
		oContext = ThisApplication.TransientObjects.CreateTranslationContext
		oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
		oOptions = ThisApplication.TransientObjects.CreateNameValueMap
		oDataMedium = ThisApplication.TransientObjects.CreateDataMedium


		'oOptions.Value("All_Color_AS_Black") = 0
		oOptions.Value("Remove_Line_Weights") = 1
		oOptions.Value("Vector_Resolution") = 400
		oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
		'oOptions.Value("Custom_Begin_Sheet") = 2
		'oOptions.Value("Custom_End_Sheet") = 4


		'get PDF target folder path
		oFolder = oPath & "\" & oAsmName & " PDF Files"

		'Check for the PDF folder and create it if it does not exist
		If Not System.IO.Directory.Exists(oFolder) Then
			System.IO.Directory.CreateDirectory(oFolder)
		End If

		'- - - - - - - - - - - - -Component Drawings - - - - - - - - - - - -
		'look at the files referenced by the assembly
		Dim oRefDocs As DocumentsEnumerator
		oRefDocs = oAsmDoc.AllReferencedDocuments
		Dim oRefDoc As Document

		'work the the drawing files for the referenced models
		'this expects that the model has a drawing of the same path and name
		For Each oRefDoc In oRefDocs
			drawingPathName = Nothing
			'check to see that the model has a drawing of the same path and name
			drawingPathName = Left(oRefDoc.FullFileName, Len(oRefDoc.FullFileName) -3) & "idw"
			If (System.IO.File.Exists(drawingPathName)) = False Then
				drawingPathName = Left(oRefDoc.FullFileName, Len(oRefDoc.FullFileName) -3) & "dwg"
				If (System.IO.File.Exists(drawingPathName)) = False Then
					Continue For
				End If
			End If

			Logger.info(oRefDoc.FullFileName)

			Dim oDrawDoc As DrawingDocument
			oDrawDoc = ThisApplication.Documents.Open(drawingPathName, False)
			oFileName = Left(oRefDoc.DisplayName, Len(oRefDoc.DisplayName) -3)

			'Set the PDF target file name
			oDataMedium.FileName = _Folder_Path & "\" & oFileName & "pdf"
			'Write out the PDF
			PDFAddIn.SaveCopyAs(oDrawDoc, oContext, oOptions, oDataMedium)
			'close the file
			oDrawDoc.Close

		Next

		drawingPathName = Nothing
		'- - - - - - - - - - - - -Top Level Drawing - - - - - - - - - - - -
		'check to see that the model has a drawing of the same path and name
		drawingPathName = ThisDoc.ChangeExtension(".idw")
		If (System.IO.File.Exists(drawingPathName)) = False Then
			drawingPathName = ThisDoc.ChangeExtension(".dwg")
			If (System.IO.File.Exists(drawingPathName)) = False Then
				drawingPathName = Nothing
			End If
		End If
				
		If Not drawingPathName = Nothing
			oAsmDrawingDoc = ThisApplication.Documents.Open(drawingPathName, False)
			
			PDFName = IO.Path.GetFileNameWithoutExtension(thisdoc.document.fullfilename)

			'Set the PDF target file name
			oDataMedium.FileName = _Folder_Path & "\" & PDFName & ".pdf"
			'Write out the PDF
			PDFAddIn.SaveCopyAs(oAsmDrawingDoc, oContext, oOptions, oDataMedium)
			'Close the top level drawing
			oAsmDrawingDoc.Close
		Else
			Logger.info("Could not find top level drawing")
			Logger.info(drawingPathName)
		End If
	End Sub




	Private Function CreatePDF(_Folder_Path As String)

		Dim bOutputfileAlreadyExists As Boolean = False
		Dim sOutFilePath As String = IO.Path.Combine(_Folder_Path, _PDF_Name & ".pdf")

		'sdet up return for a successful pdf. ret changes to FALSE if any errors occur for qualifying purposes
		Dim ret As String = sOutFilePath

		If IO.File.Exists(sOutFilePath) Then
			Try
				IO.File.Delete(sOutFilePath)
			Catch ex As Exception
				bOutputfileAlreadyExists = True
			End Try
		End If

		Dim iPageCount As Integer = GetPageCount(_Folder_Path)
		If iPageCount > 0 And bOutputfileAlreadyExists = False Then

			Dim oFiles As String() = IO.Directory.GetFiles(_Folder_Path)
			Dim oPdfDoc As New iTextSharp.text.Document()
			Dim oPdfWriter As iTextSharp.text.pdf.PdfWriter = _
			iTextSharp.text.pdf.PdfWriter.GetInstance(oPdfDoc, New IO.FileStream(sOutFilePath, IO.FileMode.Create))

			oPdfDoc.Open()

			System.Array.Sort(Of String)(oFiles)

			For i As Integer = 0 To oFiles.Length - 1
				Dim sFromFilePath As String = oFiles(i)
				Dim oFileInfo As New IO.FileInfo(sFromFilePath)
				Dim sFileType As String = "PDF"
				Dim sExt As String = PadExt(oFileInfo.Extension)

				Try
					AddPdf(sFromFilePath, oPdfDoc, oPdfWriter)
				Catch ex As Exception
					ret = "FALSE"
				End Try
			Next

			Try
				oPdfDoc.Close()
				oPdfWriter.Close()
			Catch ex As Exception

				Try
					IO.File.Delete(sOutFilePath)
				Catch ex2 As Exception
				End Try
			End Try
		End If

		Dim oFolders As String() = IO.Directory.GetDirectories(_Folder_Path)
		For i As Integer = 0 To oFolders.Length - 1
			Dim sChildFolder As String = oFolders(i)
			Dim iPos As Integer = sChildFolder.LastIndexOf("\")
			Dim sFolderName As String = sChildFolder.Substring(iPos + 1)
			CreatePDF(sChildFolder)
		Next

		Return ret

	End Function

	Private Sub AddPdf(ByVal sInFilePath As String, ByRef oPdfDoc As iTextSharp.text.Document, ByRef oPdfWriter As iTextSharp.text.pdf.PdfWriter)

		Dim oDirectContent As iTextSharp.text.pdf.PdfContentByte = oPdfWriter.DirectContent
		Dim oPdfReader As iTextSharp.text.pdf.PdfReader = New iTextSharp.text.pdf.PdfReader(sInFilePath)
		Dim iNumberOfPages As Integer = oPdfReader.NumberOfPages
		Dim iPage As Integer = 0

		Do While (iPage < iNumberOfPages)
			iPage += 1
			oPdfDoc.SetPageSize(oPdfReader.GetPageSizeWithRotation(iPage))
			oPdfDoc.NewPage()

			Dim oPdfImportedPage As iTextSharp.text.pdf.PdfImportedPage = oPdfWriter.GetImportedPage(oPdfReader, iPage)
			Dim iRotation As Integer = oPdfReader.GetPageRotation(iPage)
			If (iRotation = 90) Or (iRotation = 270) Then
				oDirectContent.AddTemplate(oPdfImportedPage, 0, -1.0F, 1.0F, 0, 0, oPdfReader.GetPageSizeWithRotation(iPage).Height)
			Else
				oDirectContent.AddTemplate(oPdfImportedPage, 1.0F, 0, 0, 1.0F, 0, 0)
			End If
		Loop

	End Sub

	Private Function PadExt(ByVal s As String) As String
		s = UCase(s)
		If s.Length > 3 Then
			s = s.Substring(1, 3)
		End If
		Return s
	End Function

	Private Function GetPageCount(ByVal _Folder_Path As String) As Integer
		Dim iRet As Integer = 0
		Dim oFiles As String() = IO.Directory.GetFiles(_Folder_Path)

		For i As Integer = 0 To oFiles.Length - 1
			Dim sFromFilePath As String = oFiles(i)
			Dim oFileInfo As New IO.FileInfo(sFromFilePath)
			Dim sFileType As String = "PDF"
			Dim sExt As String = PadExt(oFileInfo.Extension)

			iRet += 1
		Next

		Return iRet
	End Function

End Class