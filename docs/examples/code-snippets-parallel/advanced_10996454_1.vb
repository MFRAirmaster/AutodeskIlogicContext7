' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:59:25.111279

Sub Main()

   

    Dim oDoc As AssemblyDocument

     oDoc = ThisApplication.ActiveDocument

    Dim oOcc As ComponentOccurrence

     	oOcc = ThisApplication.CommandManager.Pick(SelectionFilterEnum.kAssemblyOccurrenceFilter, "Select a oOcc plate")

    Dim oMat As Matrix

     oMat = oOcc.Transformation

 

    Dim aRotAngles(2) As Double

    Call CalculateRotationAngles(oMat, aRotAngles)

   

    ' Print results

'    Dim i As Integer

'    For i = 0 To 2

      msgbox(aRotAngles(0)  & vbCrLf & aRotAngles(1) & vbCrLf & aRotAngles(2))

'    Next i

    Beep

End Sub

 

 

Sub CalculateRotationAngles(ByVal oMatrix As Inventor.Matrix,ByRef aRotAngles() As Double)

    Const PI = 3.14159265358979

    Const TODEGREES As Double = 180 / PI

 

    Dim dB As Double

    Dim dC As Double

    Dim dNumer As Double

    Dim dDenom As Double

    Dim dAcosValue As Double

       

    Dim oRotate As Inventor.Matrix

    Dim oAxis As Inventor.Vector

    Dim oCenter As Inventor.Point

   

     oRotate = ThisApplication.TransientGeometry.CreateMatrix

     oAxis = ThisApplication.TransientGeometry.CreateVector

     oCenter = ThisApplication.TransientGeometry.CreatePoint

 

    oCenter.X = 0

    oCenter.Y = 0

    oCenter.Z = 0

 

    ' Choose aRotAngles[0] about x which transforms axes[2] onto the x-z plane

    '

    dB = oMatrix.Cell(2, 3)

    dC = oMatrix.Cell(3, 3)

 

    dNumer = dC

    dDenom = Sqrt(dB * dB + dC * dC)

 

    ' Make sure we can do the division.  If not, then axes[2] is already in the x-z plane

    If (Abs(dDenom) <= 0.000001) Then

        aRotAngles(0) = 0#

    Else

        If (dNumer / dDenom >= 1#) Then

            dAcosValue = 0#

        Else

            If (dNumer / dDenom <= -1#) Then

                dAcosValue = PI

            Else

                dAcosValue = Acos(dNumer / dDenom)

            End If

        End If

   

        aRotAngles(0) = Sign(dB) * dAcosValue

        oAxis.X = 1

        oAxis.Y = 0

        oAxis.Z = 0

 

        Call oRotate.SetToRotation(aRotAngles(0), oAxis, oCenter)

        Call oMatrix.PreMultiplyBy(oRotate)

    End If

 

    '

    ' Choose aRotAngles[1] about y which transforms axes[3] onto the z axis

    '

    If (oMatrix.Cell(3, 3) >= 1#) Then

        dAcosValue = 0#

    Else

        If (oMatrix.Cell(3, 3) <= -1#) Then

            dAcosValue = PI

        Else

            dAcosValue = Acos(oMatrix.Cell(3, 3))

        End If

    End If

 

    aRotAngles(1) = Math.Sign(-oMatrix.Cell(1, 3)) * dAcosValue

    oAxis.X = 0

    oAxis.Y = 1

    oAxis.Z = 0

    Call oRotate.SetToRotation(aRotAngles(1), oAxis, oCenter)

    Call oMatrix.PreMultiplyBy(oRotate)

 

    '

    ' Choose aRotAngles[2] about z which transforms axes[0] onto the x axis

    '

    If (oMatrix.Cell(1, 1) >= 1#) Then

        dAcosValue = 0#

    Else

        If (oMatrix.Cell(1, 1) <= -1#) Then

            dAcosValue = PI

        Else

            dAcosValue = Acos(oMatrix.Cell(1, 1))

        End If

    End If

 

    aRotAngles(2) = Math.Sign(-oMatrix.Cell(2, 1)) * dAcosValue

 

    'if you want to get the result in degrees

    aRotAngles(0) = Round(aRotAngles(0) * TODEGREES,4)

    aRotAngles(1) = Round(aRotAngles(1) * TODEGREES,4)

    aRotAngles(2) = Round(aRotAngles(2) * TODEGREES,4)

End Sub

 

 

Public Function Acos(value) As Double

    Acos = Math.Atan(-value / Math.Sqrt(-value * value + 1)) + 2 * Math.Atan(1)

End Function