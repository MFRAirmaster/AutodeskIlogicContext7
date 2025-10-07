' Title: Assembly Place Component Preview
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/assembly-place-component-preview/td-p/13787147
' Category: advanced
' Scraped: 2025-10-07T13:15:16.402224

Public Shared Function PlaceOccurrence(ByVal sFilePath As String, ByVal Location_X As Double, ByVal Location_Y As Double, ByVal Location_Z As Double, ByVal Rotation_X As Double, ByVal Rotation_Y As Double, ByVal Rotation_Z As Double, ByVal Base_Location_X As Double, ByVal Base_Location_Y As Double, ByVal Base_Location_Z As Double, ByVal Base_Rotation_X As Double, ByVal Base_Rotation_Y As Double, ByVal Base_Rotation_Z As Double, Optional ByVal bGrounded As Boolean = True) As ComponentOccurrence
	Dim oResult As ComponentOccurrence

	' Set a reference to the Assembly Document
	Dim oAsmDoc As AssemblyDocument = g_inventorApplication.ActiveDocument

	' Set a reference to the Assembly Component Definition
	Dim oAsmCompDef As AssemblyComponentDefinition = oAsmDoc.ComponentDefinition

	' Create a Matrix
	Dim oOccMatrix As Matrix = MS_Matrices.CreateMatrixFromLocationAndRotation(Location_X, Location_Y, Location_Z, Rotation_X, Rotation_Y, Rotation_Z)

	' Create a Matrix
	Dim oBaseMatrix As Matrix = MS_Matrices.CreateMatrixFromLocationAndRotation(Base_Location_X, Base_Location_Y, Base_Location_Z, Base_Rotation_X, Base_Rotation_Y, Base_Rotation_Z)

	' PreMultiply the Occurrence Matrix by the Base Matrix
	oOccMatrix.PreMultiplyBy(oBaseMatrix)

	' Set a reference to the Component Occurrences
	Dim oOccurrences As ComponentOccurrences = oAsmCompDef.Occurrences

	' Place a Component Occurrence
	Dim oOcc As ComponentOccurrence = oOccurrences.Add(sFilePath, oOccMatrix)

	' Ground the Component Occurrence
	oOcc.Grounded = bGrounded

	' Set the Default View Representation if the Component Occurrence is an Assembly
	If oOcc.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then oOcc.SetDesignViewRepresentation("Default", , True)

	' Return the Component Occurrence
	oResult = oOcc

	Return oResult
End Function

Public Shared Function MoveOccurrenceToInsertionPoint(ByVal oOcc As ComponentOccurrence, ByVal sWorkPointName As String) As ComponentOccurrence
	Dim oResult As ComponentOccurrence

	' Set a reference to the Document of the Component Occurrence
	Dim oOccDoc As Document = oOcc.ReferencedDocumentDescriptor.ReferencedDocument

	' Get the Origin Point WorkPoint
	Dim oWPoint_Origin_Point As WorkPoint = oOccDoc.ComponentDefinition.WorkPoints.Item(1)

	' Create a Proxy of the Origin Point WorkPoint
	Dim oWPoint_Origin_Point_Proxy As WorkPointProxy = Nothing
	oOcc.CreateGeometryProxy(oWPoint_Origin_Point, oWPoint_Origin_Point_Proxy)

	' Check if the Occurrence should be moved to a WorkPoint other than the Origin Point WorkPoint
	If sWorkPointName <> "" AndAlso sWorkPointName <> "Center Point" Then
		Try
			' Get the Insertion Point WorkPoint
			Dim oWPoint_Insertion_Point As WorkPoint = oOccDoc.ComponentDefinition.WorkPoints.Item(sWorkPointName)

			' Create a Proxy of the Insertion Point WorkPoint 
			Dim oWPoint_Insertion_Point_Proxy As WorkPointProxy = Nothing
			oOcc.CreateGeometryProxy(oWPoint_Insertion_Point, oWPoint_Insertion_Point_Proxy)

			' Get the translation between the two WorkPoints
			Dim dTranslation_X As Double = oWPoint_Origin_Point_Proxy.Point.X - oWPoint_Insertion_Point_Proxy.Point.X
			Dim dTranslation_Y As Double = oWPoint_Origin_Point_Proxy.Point.Y - oWPoint_Insertion_Point_Proxy.Point.Y
			Dim dTranslation_Z As Double = oWPoint_Origin_Point_Proxy.Point.Z - oWPoint_Insertion_Point_Proxy.Point.Z

			' Set the translation Matrix
			Dim oTranslatedMatrix As Matrix = oOcc.Transformation
			oTranslatedMatrix.SetTranslation(g_inventorApplication.TransientGeometry.CreateVector(oWPoint_Origin_Point_Proxy.Point.X + dTranslation_X, oWPoint_Origin_Point_Proxy.Point.Y + dTranslation_Y, oWPoint_Origin_Point_Proxy.Point.Z + dTranslation_Z))

			' Set the transformation of the Component Occurrence
			oOcc.Transformation = oTranslatedMatrix
		Catch
			MsgBox("Work Point """ & sWorkPointName & """ was not found, the Occurrence was placed using the Origin Point")
		End Try
	End If

	' Return the Component Occurrence
	oResult = oOcc
	Return oResult
End Function

Public Shared Function CreateMatrixFromLocationAndRotation(ByVal Location_X As Double, ByVal Location_Y As Double, ByVal Location_Z As Double, ByVal Rotation_X As Double, ByVal Rotation_Y As Double, ByVal Rotation_Z As Double) As Matrix
	Dim oResult As Matrix

	' Set a reference to the Transient Geometry
	Dim oTransientGeometry As TransientGeometry = g_inventorApplication.TransientGeometry

	' Create Matrices
	Dim oMatrix As Matrix = oTransientGeometry.CreateMatrix
	Dim oMatrix_Z As Matrix = oTransientGeometry.CreateMatrix
	Dim oMatrix_Y As Matrix = oTransientGeometry.CreateMatrix
	Dim oMatrix_X As Matrix = oTransientGeometry.CreateMatrix

	' Set the rotation of the Matrix about the Z axis
	oMatrix_Z.SetToRotation(Rotation_Z * (Math.PI / 180), oTransientGeometry.CreateVector(0, 0, 1), oTransientGeometry.CreatePoint(0, 0, 0))

	' Set the rotation of the Matrix about the Y axis
	oMatrix_Y.SetToRotation(Rotation_Y * (Math.PI / 180), oTransientGeometry.CreateVector(0, 1, 0), oTransientGeometry.CreatePoint(0, 0, 0))

	' Set the rotation of the Matrix about the X axis
	oMatrix_X.SetToRotation(Rotation_X * (Math.PI / 180), oTransientGeometry.CreateVector(1, 0, 0), oTransientGeometry.CreatePoint(0, 0, 0))

	' PreMultiply the Matrix by the Matrices about each axis
	oMatrix.PreMultiplyBy(oMatrix_Z)
	oMatrix.PreMultiplyBy(oMatrix_Y)
	oMatrix.PreMultiplyBy(oMatrix_X)

	' Set the Translation of the matrix
	oMatrix.SetTranslation(oTransientGeometry.CreateVector(Location_X / 10, Location_Y / 10, Location_Z / 10))

	' Return the Component Occurrence
	oResult = oMatrix

	Return oResult
End Function