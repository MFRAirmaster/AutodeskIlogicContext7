' Title: iLogic (Input String was Not in a Correct Format)
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-input-string-was-not-in-a-correct-format/td-p/13784120
' Category: advanced
' Scraped: 2025-10-07T14:08:53.358264

Try
    ' Attempt to access an existing parameter
	Dim Labor_Ops_Variable As Inventor.Parameter = ThisDoc.Document.ComponentDefinition.Parameters.Item("Labor_Ops")
    ' If successful, do something with the parameter
	'Enter user defined iProperty
		customPropertySet = ThisDoc.Document.PropertySets.Item("Inventor User Defined Properties")
	
		iProperties.Value("Custom", "Labor_Ops") = Parameter("Labor_Ops")
	'Uncomment line below for test
    'MessageBox.Show("Parameter 'Labor_Ops' found with value: " & Labor_Ops_Variable.Value)

Catch
    ' If the parameter doesn't exist, create it
	Dim userParameters As UserParameters = ThisDoc.Document.ComponentDefinition.Parameters.UserParameters
	Dim newParameter As Inventor.Parameter = userParameters.AddByValue("Labor_Ops", "Saw", UnitsTypeEnum.kTextUnits)
	
	'Uncomment line below for test
    'MessageBox.Show("Parameter 'Labor_Ops' created.")
End Try
	


-----------------------------------------------------------------
' Specify your Excel file path and sheet name
    Dim excelFilePath As String = "File Path Removed for Privacy"
    Dim sheetName As String = "Inventor Form Layout"

    ' Open the Excel file
    GoExcel.Open(excelFilePath, sheetName)
	
	Dim rangeArray As Object = GoExcel.NamedRangeValue("Labor_Ops")
	
	    ' Create a list to hold the values from the Excel column
    Dim oList As ArrayList = New ArrayList()


    ' Iterate through the Excel data (assuming a single column, e.g., Column A)
    ' For each row in the rangeArray
    For iRow = rangeArray.GetLowerBound(0) To rangeArray.GetUpperBound(0)
        ' Get the value from the current row and specific column (e.g., column index 0)
		Dim cellValue As Object = rangeArray(iRow, 0) ' 0 for the first column

        ' Check if the cell is not empty before adding to the list
        If Not IsNothing(cellValue) AndAlso cellValue.ToString().Trim() <> "" Then
            oList.Add(cellValue.ToString()) ' Add the string value to the list
        End If
    Next
	
	    ' Assign the populated list to the multi-value parameter (e.g., "MyParam")
    MultiValue.List("Labor_Ops") = oList
	
	rowFound = GoExcel.FindRow(excelFilePath, sheetName, "Labor_Ops", "=", Labor_Ops)
	
	If rowFound <> -1 Then
    ' If a row is found, get the value from the "Description" column of that row
    routersValue = GoExcel.CurrentRowValue("Routers")
	'Uncomment 3 lines below to test if form is working
    'MessageBox.Show("Routers: " & routersValue)
	'Else
    'MessageBox.Show("Labor_Ops not found.")
	End If
	
	'Assigns the found cell value to the "Router_Export" parameter
	Router_Export = routersValue

	
	    ' Close the Excel file
    GoExcel.Close()