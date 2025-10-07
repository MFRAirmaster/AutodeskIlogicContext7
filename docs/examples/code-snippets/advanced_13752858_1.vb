' Title: Placing iParts into an assembly with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/placing-iparts-into-an-assembly-with-ilogic/td-p/13752858
' Category: advanced
' Scraped: 2025-10-07T12:31:34.341672

'import parts folder as string EDIT ME!
Dim partsFolder As String = "R:\Industrial Engineering\A-ME\CAD\Dolly plaques v2"

'define number of plaques EDIT ME! (could be improved by reading this value from the selected part?)
Dim numOfPlaques As Integer = 500

'define size of plaque and spacing of laser EDIT ME!
Dim plaqueLength As Integer = 280
Dim plaqueWidth As Integer = 40
Dim laserClearance As Integer = 5

'define size of sheet for nesting EDIT ME!
Dim sheetLength As Integer = 2440
Dim sheetWidth As Integer = 1220


'define x y location variables
Dim xLocation As Integer = 0
Dim yLocation As Integer = 0
Dim sheetNum As Integer = 0

'main loop for number of plaques
Dim i As Integer = 0
While i < numOfPlaques
	i += 1
	
	'debugging window text
	'Dim txt As String = String.Format("X= {0}   Y= {1}",xLocation,yLocation)
	'MessageBox.Show(txt, "Part placed:")
	
	'set name and pos for adding component
	Dim name As String = String.Format("Plaque-{0}",i)
	Dim pos = ThisDoc.Geometry.Point(xLocation,0,yLocation)

	'add Component dolley plaques v2.ipt
	Dim componentPlaque = Components.AddiPart(name, partsFolder & "\Dolley plaques v2.ipt", row := i, position :=pos, grounded := True, visible := True, appearance := Nothing)
	
	'increment collumn
	xLocation += plaqueLength + laserClearance
	
	'increment row once collumns full
	If xLocation >= sheetLength - plaqueLength + ((sheetLength + plaqueLength) * sheetNum)
		'increment row
		yLocation += plaqueWidth + laserClearance
		'go back to bottom left corner of current sheet
		xLocation = 0 + ((sheetLength + plaqueLength) * sheetNum)
	End If
	
	If yLocation >= sheetWidth - plaqueWidth
		'this loop getting entered every collumn after first sheet
		'increment sheet number
		sheetNum += 1
		'go to bottom left corner of next sheet
		xLocation = 0 + ((sheetLength + plaqueLength) * sheetNum)
		yLocation = 0
	End If
	
End While