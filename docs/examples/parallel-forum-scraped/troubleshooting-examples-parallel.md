# Parallel Troubleshooting iLogic Examples

Generated from 8 forum posts with advanced parallel scraping.

**Generated:** 2025-10-07T14:11:59.944968

---

## Using GoExcel to read a embedded spreadsheet

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/using-goexcel-to-read-a-embedded-spreadsheet/td-p/13788308](https://forums.autodesk.com/t5/inventor-programming-forum/using-goexcel-to-read-a-embedded-spreadsheet/td-p/13788308)

**Author:** harvey_craig2RCUH

**Date:** ‎08-29-2025
	
		
		02:04 AM

**Description:** I can't get GoExcel to read a cell in an embedded spreadsheet. The spreadsheet has one cell:I've embedded it using the UI:This is my local rule:MsgBox(GoExcel.CellValue("3rd Party:Embedding 1", "Sheet1", "A1")) And this is my error:I have tried both these options in the options:Both get the same object reference error.  I have also attached the drawing. Why doesn't this work? Thanks,Harvey

**Code:**

```vb
MsgBox(GoExcel.CellValue("3rd Party:Embedding 1", "Sheet1", "A1"))
```

---

## Creating a dimension from edge of circle with a face

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/creating-a-dimension-from-edge-of-circle-with-a-face/td-p/11569540](https://forums.autodesk.com/t5/inventor-programming-forum/creating-a-dimension-from-edge-of-circle-with-a-face/td-p/11569540)

**Author:** phucminhnguyen.76

**Date:** ‎11-22-2022
	
		
		01:42 AM

**Description:** Hi everyone, I started to use iLogic and create Rules in my projects recently.I would like to create a dimension like my attached image. Does someone has the iLogic experience to help me?Thanks!

**Code:**

```vb
' Get sheet and view
Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

' Get named entities
Dim VFB_Face_Left = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Left", PointIntentEnum.kEndPointIntent)
Dim VFB_Face_Bottom = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Bottom", PointIntentEnum.kEndPointIntent)
Dim VRB_Face_Cylinder_Right = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularRightPointIntent)
Dim VRB_Face_Cylinder_Top = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularTopPointIntent)

' Declare access to drawing dimensions
Dim genDims = Sheet_1.DrawingDimensions.GeneralDimensions

' Add the dimensions
Try
	genDims.AddLinear("Dim01", VIEW1.SheetPoint(0.5, -0.1875), VFB_Face_Left, VRB_Face_Cylinder_Right, DimensionTypeEnum.kHorizontalDimensionType)
	genDims.AddLinear("Dim02", VIEW1.SheetPoint(-0.03125, 0), VFB_Face_Bottom, VRB_Face_Cylinder_Top, DimensionTypeEnum.kVerticalDimensionType)
Catch ex As Exception
	MsgBox("Error: ", ex.Message )
End Try
```

```vb
' Get sheet and view
Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

' Get named entities
Dim VFB_Face_Left = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Left", PointIntentEnum.kEndPointIntent)
Dim VFB_Face_Bottom = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Bottom", PointIntentEnum.kEndPointIntent)
Dim VRB_Face_Cylinder_Right = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularRightPointIntent)
Dim VRB_Face_Cylinder_Top = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularTopPointIntent)

' Declare access to drawing dimensions
Dim genDims = Sheet_1.DrawingDimensions.GeneralDimensions

' Add the dimensions
Try
	genDims.AddLinear("Dim01", VIEW1.SheetPoint(0.5, -0.1875), VFB_Face_Left, VRB_Face_Cylinder_Right, DimensionTypeEnum.kHorizontalDimensionType)
	genDims.AddLinear("Dim02", VIEW1.SheetPoint(-0.03125, 0), VFB_Face_Bottom, VRB_Face_Cylinder_Top, DimensionTypeEnum.kVerticalDimensionType)
Catch ex As Exception
	MsgBox("Error: ", ex.Message )
End Try
```

---

## iLogic code to zip files suddenly stopped working

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-zip-files-suddenly-stopped-working/td-p/13180134](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-zip-files-suddenly-stopped-working/td-p/13180134)

**Author:** b.graaf

**Date:** ‎11-28-2024
	
		
		07:04 AM

**Description:** Hi! I have an iLogic code which generate several stepfiles and add them into a new created compressed folder (ZIP).I am using it in several Inventor versions and we are now on the latest Inventor professional 2025. I have used this iLogic rule about a year ago for the last time and now it suddenly does not work anymore!I did not change anything in the code. Is there something changed within iLogic, or maybe in .net? I am using this line:  System.IO.Compression.ZipFile.CreateFromDirectory(_export...

**Code:**

```vb
System.IO.Compression.ZipFile.CreateFromDirectory(_export_Location, _ZipFileName)
```

```vb
System.IO.Compression.ZipFile.CreateFromDirectory(_export_Location, _ZipFileName)
```

```vb
System.IO.Compression.FileSystem
```

```vb
AddReference "System.IO.Compression"
AddReference "System.IO.Compression.FileSystem"
```

```vb
AddReference "System.IO.Compression"
AddReference "System.IO.Compression.FileSystem"
```

```vb
AddReference "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.IO.Compression.FileSystem.dll"
```

```vb
AddReference "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.IO.Compression.FileSystem.dll"
```

---

## Excel Error with iLogic

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/excel-error-with-ilogic/td-p/13752877](https://forums.autodesk.com/t5/inventor-programming-forum/excel-error-with-ilogic/td-p/13752877)

**Author:** layochim

**Date:** ‎08-04-2025
	
		
		08:01 AM

**Description:** I'm sure this is a common error, but I'm unable to fix it.  I've got a Pricing rule below that works the first time then when i modify a parameter i get the following.  Anyone have a solution for this?My code is:Dim Filename As String = "Straight Conveyor Pricing.xlsx"
Dim SheetName As String = "Conv costing"
GoExcel.Open(Filename, SheetName)

GoExcel.CellValue("B9") = BeltWidth
GoExcel.CellValue("B10") = BeltLength
GoExcel.CellValue("B11") = Infeed_End
GoExcel.CellValue("B12") = Discharge_End
G...

**Code:**

```vb
Dim Filename As String = "Straight Conveyor Pricing.xlsx"
Dim SheetName As String = "Conv costing"
GoExcel.Open(Filename, SheetName)

GoExcel.CellValue("B9") = BeltWidth
GoExcel.CellValue("B10") = BeltLength
GoExcel.CellValue("B11") = Infeed_End
GoExcel.CellValue("B12") = Discharge_End
GoExcel.CellValue("B13") = Legs
GoExcel.CellValue("B14") = Feet
GoExcel.CellValue("B15") = Belting
GoExcel.CellValue("B16") = BeltingType
GoExcel.CellValue("B17") = BeltColor
GoExcel.CellValue("B18") = Flights
GoExcel.CellValue("B19") = Flight_Spacing
GoExcel.CellValue("B20") = Construction
GoExcel.CellValue("B21") = FlangeHeight
GoExcel.CellValue("B22") = Finish
GoExcel.CellValue("B24") = Drive_Style
GoExcel.CellValue("B25") = Drive_Position
GoExcel.CellValue("B26") = MotorVoltage
GoExcel.CellValue("B27") = Drive_Speed
GoExcel.CellValue("B28") = Motor
GoExcel.CellValue("B29") = Gearbox
GoExcel.CellValue("B30") = AdjustableGuides
GoExcel.CellValue("B31") = BeltLifters
GoExcel.CellValue("B32") = DripPan
GoExcel.CellValue("B33") = BeltScraper
GoExcel.CellValue("B34") = SprayBar
GoExcel.CellValue("B35") = GearboxGuards

Price = GoExcel.CellValue("B6")

GoExcel.Save
GoExcel.Close
```

```vb
Dim Filename As String = "C:/Folder/Another Folder/Straight Conveyor Pricing.xlsx"
```

---

## Creating a dimension from edge of circle with a face

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/creating-a-dimension-from-edge-of-circle-with-a-face/td-p/11569540#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/creating-a-dimension-from-edge-of-circle-with-a-face/td-p/11569540#messageview_0)

**Author:** phucminhnguyen.76

**Date:** ‎11-22-2022
	
		
		01:42 AM

**Description:** Hi everyone, I started to use iLogic and create Rules in my projects recently.I would like to create a dimension like my attached image. Does someone has the iLogic experience to help me?Thanks!

**Code:**

```vb
' Get sheet and view
Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

' Get named entities
Dim VFB_Face_Left = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Left", PointIntentEnum.kEndPointIntent)
Dim VFB_Face_Bottom = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Bottom", PointIntentEnum.kEndPointIntent)
Dim VRB_Face_Cylinder_Right = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularRightPointIntent)
Dim VRB_Face_Cylinder_Top = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularTopPointIntent)

' Declare access to drawing dimensions
Dim genDims = Sheet_1.DrawingDimensions.GeneralDimensions

' Add the dimensions
Try
	genDims.AddLinear("Dim01", VIEW1.SheetPoint(0.5, -0.1875), VFB_Face_Left, VRB_Face_Cylinder_Right, DimensionTypeEnum.kHorizontalDimensionType)
	genDims.AddLinear("Dim02", VIEW1.SheetPoint(-0.03125, 0), VFB_Face_Bottom, VRB_Face_Cylinder_Top, DimensionTypeEnum.kVerticalDimensionType)
Catch ex As Exception
	MsgBox("Error: ", ex.Message )
End Try
```

```vb
' Get sheet and view
Dim Sheet_1 = ThisDrawing.Sheets.ItemByName("Sheet:1")
Dim VIEW1 = Sheet_1.DrawingViews.ItemByName("VIEW1")

' Get named entities
Dim VFB_Face_Left = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Left", PointIntentEnum.kEndPointIntent)
Dim VFB_Face_Bottom = VIEW1.GetIntent("Vert_Flat_Bar_01", "Face_Bottom", PointIntentEnum.kEndPointIntent)
Dim VRB_Face_Cylinder_Right = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularRightPointIntent)
Dim VRB_Face_Cylinder_Top = VIEW1.GetIntent("Vert_Rnd_Bar", "Face_Cylinder", PointIntentEnum.kCircularTopPointIntent)

' Declare access to drawing dimensions
Dim genDims = Sheet_1.DrawingDimensions.GeneralDimensions

' Add the dimensions
Try
	genDims.AddLinear("Dim01", VIEW1.SheetPoint(0.5, -0.1875), VFB_Face_Left, VRB_Face_Cylinder_Right, DimensionTypeEnum.kHorizontalDimensionType)
	genDims.AddLinear("Dim02", VIEW1.SheetPoint(-0.03125, 0), VFB_Face_Bottom, VRB_Face_Cylinder_Top, DimensionTypeEnum.kVerticalDimensionType)
Catch ex As Exception
	MsgBox("Error: ", ex.Message )
End Try
```

---

## Overridden Balloons

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/overridden-balloons/td-p/13699659#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/overridden-balloons/td-p/13699659#messageview_0)

**Author:** gerardo_arellanoPYCV8

**Date:** ‎06-26-2025
	
		
		07:21 AM

**Description:** Hello, I've been working with a macro that iterates through the balloons in a drawing with the intention of compare the balloons to the BOM.  I've tried several ways and did not work.The main issue is getting the information of overridden balloons like this format = (item / qty).I want to get the information of the qty overridden value, and I didn't have luck.Is always an error like this: Object doesn't support this property or method.reading more about it could be that my VBA project is not cor...

**Code:**

```vb
Select Case oBalloon.Type
```

```vb
Select Case oBalloon.Type
```

```vb
Select Case oBalloon.GetBalloonType
```

```vb
Select Case oBalloon.GetBalloonType
```

```vb
Dim balloonType As BalloonTypeEnum
Dim balloonTypeData As Variant
Call oBalloon.GetBalloonType(balloonType, balloonTypeData)
Debug.Print "Balloon type = " & balloonType
```

```vb
Dim balloonType As BalloonTypeEnum
Dim balloonTypeData As Variant
Call oBalloon.GetBalloonType(balloonType, balloonTypeData)
Debug.Print "Balloon type = " & balloonType
```

---

## iLogic code to zip files suddenly stopped working

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-zip-files-suddenly-stopped-working/td-p/13180134#messageview_0](https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-code-to-zip-files-suddenly-stopped-working/td-p/13180134#messageview_0)

**Author:** b.graaf

**Date:** ‎11-28-2024
	
		
		07:04 AM

**Description:** Hi! I have an iLogic code which generate several stepfiles and add them into a new created compressed folder (ZIP).I am using it in several Inventor versions and we are now on the latest Inventor professional 2025. I have used this iLogic rule about a year ago for the last time and now it suddenly does not work anymore!I did not change anything in the code. Is there something changed within iLogic, or maybe in .net? I am using this line:  System.IO.Compression.ZipFile.CreateFromDirectory(_export...

**Code:**

```vb
System.IO.Compression.ZipFile.CreateFromDirectory(_export_Location, _ZipFileName)
```

```vb
System.IO.Compression.ZipFile.CreateFromDirectory(_export_Location, _ZipFileName)
```

```vb
System.IO.Compression.FileSystem
```

```vb
AddReference "System.IO.Compression"
AddReference "System.IO.Compression.FileSystem"
```

```vb
AddReference "System.IO.Compression"
AddReference "System.IO.Compression.FileSystem"
```

```vb
AddReference "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.IO.Compression.FileSystem.dll"
```

```vb
AddReference "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.8\System.IO.Compression.FileSystem.dll"
```

---

## Overridden Balloons

**Source:** [https://forums.autodesk.com/t5/inventor-programming-forum/overridden-balloons/td-p/13699659](https://forums.autodesk.com/t5/inventor-programming-forum/overridden-balloons/td-p/13699659)

**Author:** gerardo_arellanoPYCV8

**Date:** ‎06-26-2025
	
		
		07:21 AM

**Description:** Hello, I've been working with a macro that iterates through the balloons in a drawing with the intention of compare the balloons to the BOM.  I've tried several ways and did not work.The main issue is getting the information of overridden balloons like this format = (item / qty).I want to get the information of the qty overridden value, and I didn't have luck.Is always an error like this: Object doesn't support this property or method.reading more about it could be that my VBA project is not cor...

**Code:**

```vb
Select Case oBalloon.Type
```

```vb
Select Case oBalloon.Type
```

```vb
Select Case oBalloon.GetBalloonType
```

```vb
Select Case oBalloon.GetBalloonType
```

```vb
Dim balloonType As BalloonTypeEnum
Dim balloonTypeData As Variant
Call oBalloon.GetBalloonType(balloonType, balloonTypeData)
Debug.Print "Balloon type = " & balloonType
```

```vb
Dim balloonType As BalloonTypeEnum
Dim balloonTypeData As Variant
Call oBalloon.GetBalloonType(balloonType, balloonTypeData)
Debug.Print "Balloon type = " & balloonType
```

---

