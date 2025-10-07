' Title: Issue with Vault revision table
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/issue-with-vault-revision-table/td-p/13767703
' Category: advanced
' Scraped: 2025-10-07T14:03:13.194110

Sub Main()
    Dim app As Application = ThisApplication
    Dim oDoc As DrawingDocument = app.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    If oDoc.DocumentType <> kDrawingDocumentObject Then Exit Sub

    Dim offsetX As Double = 4*2.54
    Dim offsetY As Double = 4.75*2.54

    For Each oSheet As Sheet In oDoc.Sheets
        ' Vérifier qu'il y a au moins un tableau
        If oSheet.RevisionTables.Count = 0 Then Continue For

        Dim oRevTable As RevisionTable = oSheet.RevisionTables.Item(1)

        ' Vérification préalable : uniquement si la feuille est active
        If oSheet.Name <> oDoc.ActiveSheet.Name Then
            ' Skip ou copier les valeurs du tableau pour éviter le crash
            Continue For
        End If

        ' Coin inférieur gauche du cartouche
        Dim cartouchePos As Point2d = oSheet.TitleBlock.Position
        Dim targetX As Double = cartouchePos.X + offsetX
        Dim targetY As Double = cartouchePos.Y + offsetY

        ' Boîte englobante du tableau
        Dim box As Box2d = oRevTable.RangeBox
        Dim anchor As Point2d = oRevTable.Position

        Dim vX As Double = box.MinPoint.X - anchor.X
        Dim vY As Double = box.MinPoint.Y - anchor.Y

        Dim TG As TransientGeometry = app.TransientGeometry
        Dim newAnchor As Point2d = TG.CreatePoint2d(targetX - vX, targetY - vY)

        oRevTable.Position = newAnchor
    Next
End Sub