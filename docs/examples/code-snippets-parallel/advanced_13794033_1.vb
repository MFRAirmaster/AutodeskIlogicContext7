' Title: Automation for sheet usage
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/automation-for-sheet-usage/td-p/13794033
' Category: advanced
' Scraped: 2025-10-07T14:05:19.303496

Sub ModifyPageScale()
		
		Dim Version As String = iProperties.Value("Project", "Revision Number")
		Dim JobNo As String = iProperties.Value("Project", "Project")
	
	Try	
		Dim oDoc As DrawingDocument = ThisApplication.ActiveDocument
		Dim oTG As TransientGeometry = ThisApplication.TransientGeometry
		Dim oScale As Double = .25 / 12
		
		Dim oPage1 As Sheet = oDoc.Sheets.Item(1)
	
		Dim oPageOneView1 As DrawingView
		Dim oPageOneView2 As DrawingView
		Dim oPageOneView3 As DrawingView 
		Dim oPageOneView4 As DrawingView
	
		Try
			oPageOneView1 = oPage1.DrawingViews.Item(4) 'VIEW AT 135 DEG - View 1
			oPageOneView2 = oPage1.DrawingViews.Item(3) 'VIEW AT 45 DEG - View 2
			oPageOneView3 = oPage1.DrawingViews.Item(2) 'VIEW AT 315 DEG - View 3
			oPageOneView4 = oPage1.DrawingViews.Item(1) 'VIEW AT 225 DEG - View 4
		Catch ex As Exception
			MessageBox.Show("Problem at document creation, assembly, part, or otherwise was not selected correctly at start up." * vbCrLf & "Please close this document and restart" & vbCrLf & "Exitting startup program.")
			Exit Sub
		End Try
		
		oPageOneView1.Scale = oScale
		oPageOneView2.Scale = oScale
		oPageOneView3.Scale = oScale
		oPageOneView4.Scale = oScale
		
		'MessageBox.Show("Reset scale for tesing.")
		
		Dim ViewStyle As DrawingViewStyleEnum = DrawingViewStyleEnum.kHiddenLineDrawingViewStyle
		
		Dim Box1Point As Point2d = oTG.CreatePoint2d(.635, .635)
		Dim Box2Point As Point2d = oTG.CreatePoint2d(33, 13)
		
		Dim oPage1Box As Box2d = oTG.CreateBox2d()
		oPage1Box.MinPoint = Box1Point
		oPage1Box.MaxPoint = Box2Point
		
		Dim oPage1BoxWidth As Double = Box2Point.X - Box1Point.X
		Dim oPage1BoxHeight As Double = Box2Point.Y - Box1Point.Y
		
		'MessageBox.Show(oPage1BoxWidth.ToString & " " & oPage1BoxHeight.ToString) 'This checks the size of the x and y of the box object
		
		Dim oBox1YHeight As Double = oPage1BoxHeight / 2 + .635 'This variable represents the distance from center to bottom of page for page one view placement
		Dim oBox1XWidthStart As Double = (oPage1BoxWidth / 4) / 2 + .635 'This variable represents the distance from left edge of sheet to center start point for first page 1st view placement
		Dim oBox1XWidthCenter As Double = (oPage1BoxWidth / 4) / 2 'This value represents the distance on the x axis for view spacing. 
		
		'Messagebox.Show("The first value should match 6.945 " & oBox1Yheight.ToString & " This second value should match 4.68 " & oBox1XWidthStart.ToString & " The thrid value should match 4.0456 " & oBox1XWidthCenter.ToString) 'This verifies the correct distance values mathmatically. 
		
		'Below establishes the point objects that each view on page one will be attached to upon runtime. 
		Dim PageOneViewOne As Point2d = oTG.CreatePoint2d(oBox1XWidthStart, oBox1YHeight)
		Dim PageOneViewTwo As Point2d = oTG.CreatePoint2d(oBox1XWidthStart + oBox1XWidthCenter * 2, oBox1YHeight)
		Dim PageOneViewThree As Point2d = oTG.CreatePoint2d(oBox1XWidthStart + oBox1XWidthCenter * 4, oBox1YHeight)
		Dim PageOneviewFour As Point2d = oTG.CreatePoint2d(oBox1XWidthStart + oBox1XWidthCenter * 6, oBox1YHeight)
		
		'Start the scaling of each view in sheet one to be smaller than the quadrant of the bounding box. 
		'Using height, and width of the views, we can compare that to a variable established to represent the size of the quadrant they're within. 
		Dim Page1MaxViewWidth As Double = oPage1BoxWidth / 4 ' Should equal 8.091
		Dim Page1MaxViewHeight As Double = oPage1BoxHeight 'Should be the same as the total height of the box, no reason to equate it. 
		
		i = 0
		oScale = oPageOneView1.Scale
		
		'The section below is checking the view size against the created box2d zones in the program. If the height or the width are too large, we scale by 94% and check the values again until the view size is lower than the box2d object
		Do While oPageOneView1.Width > Page1MaxViewWidth Or oPageOneView1.Height > Page1MaxViewHeight
			'
			'MessageBox.Show("Page one width: " & oPageOneView1.Width.ToString & " Page one height: " & oPageOneView1.Height.ToString & vbCrLf & " Page 1 width maximum: " & Page1MaxViewWidth & " Page 1 height maximum: " & Page1MaxViewHeight)
			oScale = oScale * .96
			oPageOneView1.Scale = oScale
			oPageOneView2.Scale = oScale
			oPageOneView3.Scale = oScale
			oPageOneView4.Scale = oScale
		
			'Logger.Info("Loop " & i & " - oScale:" & oScale & " Viewscale: " & oPageOneView1.Scale)
			'Logger.Info("   " & oPageOneView1.Width)
			'Logger.Info("   " & oPageOneView1.Height)
			'Logger.Info("   View Width: " & oPageOneView1.Width)
			'Logger.Info("   View Height: " & oPageOneView1.Height)
		
			'ActiveSheet.NativeEntity.Update
			'ThisDoc.Document.Update
		
			'endless loop killer just in case. This is a file within the C drive that exists to create a stop barrior between an endless loop and restarting the program. 
			'If the file is not found, or has it's name changed. It will enter the if block and exit the sub immediately before the need to crash the software. 
			'oFile = "C:\Temp\Endless Loop Saftey File.txt"
			'If IO.File.Exists("C:\Temp\Endless Loop Saftey File.txt") = False Then
			'	MsgBox("The rule called '" & iLogicVb.RuleName & "' stopped because text file doesn't exist:" & vbLf & oFile, , "iLogic")
			'	Exit Sub
			'End If
			
			
			i = i + 1
		Loop
	Catch ex As Exception
		MessageBox.Show("Please copy the following message into ""Bug reporting"" on teams:" & vbCrLf & vbCrLf & "Drawing template - Updating symbol library." & vbCrLf & _
		"Version: " & Version & vbCrLf & "Job no: " & JobNo & vbCrLf & "Unable to scale page 1 views, please check placement and size. ") 
	End Try
		
End Sub