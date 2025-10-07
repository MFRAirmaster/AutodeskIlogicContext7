' Title: Auto constrain in current position with Origin Planes flush or mate
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/auto-constrain-in-current-position-with-origin-planes-flush-or/td-p/10996454#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:33:58.654638

Sub main()
    Dim assemblydoc As AssemblyDocument
     assemblydoc = ThisApplication.ActiveDocument

    ' Get the occurrences in the select set.
    Dim occurrenceList As New Collection
    Dim entity As Object
    For Each entity In assemblydoc.SelectSet
        If TypeOf entity Is ComponentOccurrence Then
            occurrenceList.Add(entity)
        End If
    Next

    If occurrenceList.Count < 2 Then
        MsgBox("At least two occurrences must be selected.")
        Exit Sub
    End If

    ' This assumes the first selected occurrence is the "base"
    ' and will constrain the base workplanes of all the other parts
    ' to the base workplanes of the first part. If there are
    ' constraints on the other they end up being over constrained.

    ' Get the planes from the base part and create proxies for them.
    Dim baseOccurrence As ComponentOccurrence
     baseOccurrence = occurrenceList.Item(1)

    Dim BaseXY As WorkPlane
    Dim BaseYZ As WorkPlane
    Dim BaseXZ As WorkPlane
    Call GetPlanes(baseOccurrence, BaseXY, BaseYZ, BaseXZ)

    Dim constraints As AssemblyConstraints
     constraints = assemblydoc.ComponentDefinition.Constraints

    ' Iterate through the other occurrences
    Dim i As Integer
    For i = 2 To occurrenceList.Count
        Dim thisOcc As ComponentOccurrence
         thisOcc = occurrenceList.Item(i)

        ' Move it to the base occurrence so that if the base is
        ' not fully constrained it shouldn't move when the flush
        ' constraints are added.
        thisOcc.Transformation = baseOccurrence.Transformation

        ' Get the planes from the occurrence
        Dim occPlaneXY As WorkPlane
        Dim occPlaneYZ As WorkPlane
        Dim occPlaneXZ As WorkPlane
        Call GetPlanes(thisOcc, occPlaneXY, occPlaneYZ, occPlaneXZ)

        ' Add the flush constraints.
        Call constraints.AddFlushConstraint(BaseXY, occPlaneXY, 0)
        Call constraints.AddFlushConstraint(BaseYZ, occPlaneYZ, 0)
        Call constraints.AddFlushConstraint(BaseXZ, occPlaneXZ, 0)
    Next
End Sub

' Utility function used by the AlignOccurrencesWithConstraints macro.
' Given an occurrence it returns the base work planes that are in
' the part or assembly the occurrence references.  It gets the
' proxies for the planes since it needs the work planes in the
' context of the assembly and not in the part or assembly document
' where they actually exist.
Private Sub GetPlanes(ByVal Occurrence As ComponentOccurrence, _ 
                      ByRef BaseXY As WorkPlane, _ 
                      ByRef BaseYZ As WorkPlane, _ 
                      ByRef BaseXZ As WorkPlane)
    ' Get the work planes from the definition of the occurrence.
    ' These will be in the context of the part or subassembly, not 
    ' the top-level assembly, which is what we need to return.
     BaseXY = Occurrence.Definition.WorkPlanes.Item(3)
     BaseYZ = Occurrence.Definition.WorkPlanes.Item(1)
     BaseXZ = Occurrence.Definition.WorkPlanes.Item(2)

    ' Create proxies for these planes.  This will act as the work
    ' plane in the context of the top-level assembly.
    Call Occurrence.CreateGeometryProxy(BaseXY, BaseXY)
    Call Occurrence.CreateGeometryProxy(BaseYZ, BaseYZ)
    Call Occurrence.CreateGeometryProxy(BaseXZ, BaseXZ)
End Sub