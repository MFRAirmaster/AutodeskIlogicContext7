' Title: iLogic for hide/show every possible object
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-for-hide-show-every-possible-object/td-p/13761985
' Category: advanced
' Scraped: 2025-10-07T13:30:44.523203

Sub Main()

' iLogic: Sichtbarkeit umschalten je nach Dokumenttyp und Auswahl

Dim oDoc As Document = ThisApplication.ActiveDocument
Dim oSelSet As SelectSet = oDoc.SelectSet

If oSelSet.Count = 0 Then
    MsgBox("Selektiere zuerst etwas (Bauteil, Ebene, Skizze...)", vbExclamation)
    Return
End If

' Ursprüngliche Auswahl sichern
Dim savedSelection As New List(Of Object)
For Each oObj In oSelSet
    savedSelection.Add(oObj)
Next

For Each oObj In oSelSet
     Call ToggleVisibility(oObj)
Next

' Auswahl wiederherstellen
oSelSet.Clear()
For Each oObj In savedSelection
    Try
        oSelSet.Select(oObj)
    Catch
        ' Falls Objekt nicht mehr auswählbar ist (z.B. ausgeblendet)
    End Try
Next


End Sub

'--------------------------------------
' Hilfsfunktion: Sichtbarkeit umschalten
'--------------------------------------
Sub ToggleVisibility(oObj As Object)
    Try
        ' Falls das Objekt die Property "Visible" hat
        Dim t As Type = oObj.GetType()
        Dim p = t.GetProperty("Visible")
        If Not p Is Nothing Then
            Dim current As Boolean = CBool(p.GetValue(oObj, Nothing))
            p.SetValue(oObj, Not current, Nothing)
            Exit Sub
        End If
		
        ' Spezielle Fälle behandeln:
        If TypeOf oObj Is ComponentOccurrence Then
            oObj.Visible = Not oObj.Visible

        ' Skizzen im Part
        ElseIf TypeOf oObj Is PlanarSketch Then
            oObj.Visible = Not oObj.Visible
        ElseIf TypeOf oObj Is Sketch3D Then
            oObj.Visible = Not oObj.Visible


        ' Skizzen in der Baugruppe (Proxy)
        ElseIf TypeOf oObj Is PlanarSketchProxy Then
		    Dim oProxy As PlanarSketchProxy = oObj
		    Dim oNative As PlanarSketch = oProxy.NativeObject
		    oNative.Visible = Not oNative.Visible
		    
		    ' Occurrence im Assembly holen und sicherstellen, dass Skizzen angezeigt werden
		    Dim occ As ComponentOccurrence = oProxy.ContainingOccurrence
		    If Not occ Is Nothing Then
		        If oNative.Visible Then
		            occ.Visible = True ' Stellt sicher, dass Komponente sichtbar ist
		            occ.Parent.ShowComponentSketches = True ' Skizzenanzeige aktivieren
		        End If
		    End If
        ElseIf TypeOf oObj Is Sketch3DProxy Then
		    Dim oProxy As Sketch3DProxy = oObj
		    Dim oNative As Sketch3D = oProxy.NativeObject
		    oNative.Visible = Not oNative.Visible
		
		    Dim occ As ComponentOccurrence = oProxy.ContainingOccurrence
		    If Not occ Is Nothing Then
		        If oNative.Visible Then
		            occ.Visible = True
		            occ.Parent.ShowComponentSketches = True
		        End If
		    End If


		' Arbeitsebenen
		ElseIf TypeOf oObj Is WorkPlane Then
		    oObj.Visible = Not oObj.Visible
		ElseIf TypeOf oObj Is WorkPlaneProxy Then
		    Dim oProxy As WorkPlaneProxy = oObj
		    Dim oNative As WorkPlane = oProxy.NativeObject
		    oNative.Visible = Not oNative.Visible
		    ' Sicherstellen, dass Arbeitsebenen in der Baugruppe angezeigt werden
		    Dim occ As ComponentOccurrence = oProxy.ContainingOccurrence
		    If Not occ Is Nothing AndAlso oNative.Visible Then
		        occ.Visible = True
		        occ.Parent.ShowComponentWorkPlanes = True
		    End If
		
		' Arbeitsachsen
		ElseIf TypeOf oObj Is WorkAxis Then
		    oObj.Visible = Not oObj.Visible
		ElseIf TypeOf oObj Is WorkAxisProxy Then
		    Dim oProxy As WorkAxisProxy = oObj
		    Dim oNative As WorkAxis = oProxy.NativeObject
		    oNative.Visible = Not oNative.Visible
		    Dim occ As ComponentOccurrence = oProxy.ContainingOccurrence
		    If Not occ Is Nothing AndAlso oNative.Visible Then
		        occ.Visible = True
		        occ.Parent.ShowComponentWorkAxes = True
		    End If
		
		' Arbeitspunkte
		ElseIf TypeOf oObj Is WorkPoint Then
		    oObj.Visible = Not oObj.Visible
		ElseIf TypeOf oObj Is WorkPointProxy Then
		    Dim oProxy As WorkPointProxy = oObj
		    Dim oNative As WorkPoint = oProxy.NativeObject
		    oNative.Visible = Not oNative.Visible
		    Dim occ As ComponentOccurrence = oProxy.ContainingOccurrence
		    If Not occ Is Nothing AndAlso oNative.Visible Then
		        occ.Visible = True
		        occ.Parent.ShowComponentWorkPoints = True
		    End If

		' Komponentenmuster Rund
		ElseIf TypeOf oObj Is CircularOccurrencePattern Then
   			 oObj.Visible = Not oObj.Visible

		' Komponentenmuster Rechteckig
		ElseIf TypeOf oObj Is RectangularOccurrencePattern Then
		    oObj.Visible = Not oObj.Visible
		
		' Flächenkörper
		ElseIf TypeOf oObj Is SurfaceBody Then
		    oObj.Visible = Not oObj.Visible

        ' Zeichnungselemente
        ElseIf TypeOf oObj Is DrawingCurveSegment Then
            oObj.Visible = Not oObj.Visible

        Else
             'Optional: Debug-Ausgabe für unbekannte Typen
			 MsgBox("Kein Sichtbarkeits-Flag für: " & TypeName(oObj) & vbCrLf & "Voller Name: " & oObj.GetType().FullName)
		End If

    Catch ex As Exception
        MsgBox("Fehler bei: " & oObj.ToString & vbCrLf & TypeName(oObj) & vbCrLf & ex.Message)
    End Try
End Sub