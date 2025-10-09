' Title: Issue with Vault revision table
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/issue-with-vault-revision-table/td-p/13767703#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:07:38.123295

Sub Main()
    Dim app As Application = ThisApplication
    Dim oDoc As DrawingDocument = app.ActiveDocument
    If oDoc Is Nothing Then Exit Sub
    If oDoc.DocumentType <> kDrawingDocumentObject Then Exit Sub

    Dim oSheet As Sheet = oDoc.ActiveSheet
    If oSheet.RevisionTables.Count = 0 Then Exit Sub

    Dim oRevTable As RevisionTable = oSheet.RevisionTables.Item(1)

    ' Coin inférieur gauche du cartouche
    Dim cartouchePos As Point2d = oSheet.TitleBlock.Position

    ' Décalage depuis le coin inférieur gauche du cartouche (en cm)
    Dim offsetX As Double = 4*2.54   ' ex.: 2 cm depuis le bord gauche du cartouche
    Dim offsetY As Double = 4.75*2.54   ' ex.: 0.5 cm depuis le bord bas du cartouche

    ' Point cible bas-gauche du tableau sur la feuille
    Dim targetX As Double = cartouchePos.X + offsetX
    Dim targetY As Double = cartouchePos.Y + offsetY

    ' Boîte englobante actuelle du tableau
    Dim box As Box2d = oRevTable.RangeBox
    Dim anchor As Point2d = oRevTable.Position

    ' Vecteur du point d’ancrage actuel vers le bas-gauche réel
    Dim vX As Double = box.MinPoint.X - anchor.X
    Dim vY As Double = box.MinPoint.Y - anchor.Y

    ' Calculer le nouvel ancrage pour aligner le bas-gauche sur la cible
    Dim TG As TransientGeometry = app.TransientGeometry
    Dim newAnchor As Point2d = TG.CreatePoint2d(targetX - vX, targetY - vY)

    ' Positionner le tableau
    oRevTable.Position = newAnchor
End Sub