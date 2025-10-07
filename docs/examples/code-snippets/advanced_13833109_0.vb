' Title: Ataching leader note to sketched symbol in drawing
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ataching-leader-note-to-sketched-symbol-in-drawing/td-p/13833109
' Category: advanced
' Scraped: 2025-10-07T13:27:39.447602

Public Sub AddLeaderNotetestxx()
    Dim trg As TransientGeometry
    Set trg = ThisApplication.TransientGeometry
    
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.activeDocument
    
    Dim sk As SketchedSymbol
   
    Dim oActiveSheet As Sheet
    Set oActiveSheet = oDrawDoc.ActiveSheet

    Set sk = oActiveSheet.SketchedSymbols(1).leader.Parent

    Dim oMidPoint As Point2d
    Set oMidPoint = sk.Position

    Dim oLeaderPoints As ObjectCollection
  
    Set oLeaderPoints = ThisApplication.TransientObjects.CreateObjectCollection
    
    Dim oGeometryIntent As GeometryIntent
    Call oLeaderPoints.Add(trg.CreatePoint2d(oMidPoint.x + 5, oMidPoint.Y + 5))
    Call oLeaderPoints.Add(trg.CreatePoint2d(sk.Position.x, sk.Position.Y + 1.1))
    
    Set oGeometryIntent = oActiveSheet.CreateGeometryIntent(sk.leader.Parent, trg.CreatePoint2d(sk.Position.x, sk.Position.Y + 1.1))
    
    Call oLeaderPoints.Add(oGeometryIntent)

    Dim sText As String
    sText = "API Leader Note"

    Dim oLeaderNote As LeaderNote
    Set oLeaderNote = oActiveSheet.DrawingNotes.LeaderNotes.Add(oLeaderPoints, sText)
    
End Sub