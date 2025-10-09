' Title: SelectionFilterEnum for holes
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/selectionfilterenum-for-holes/td-p/12017592
' Category: advanced
' Scraped: 2025-10-09T08:56:26.433701

Public Class ThisRule

    Sub Main()


        Dim selector As New Selector(ThisApplication)
        selector.Pick()

        Dim doc As PartDocument = ThisDoc.Document
        Dim def As PartComponentDefinition = doc.ComponentDefinition

        Dim sketch As PlanarSketch = def.Sketches.Add(selector.SelectedFace)

        Dim point2d As Point2d = sketch.ModelToSketchSpace(selector.ModelPosition)

        Dim circleRadius = 0.5 ' in Cm!
        sketch.SketchCircles.AddByCenterRadius(point2d, circleRadius)
        Dim profile = sketch.Profiles.AddForSolid()


        Dim exturdDef = def.Features.ExtrudeFeatures.CreateExtrudeDefinition(profile, PartFeatureOperationEnum.kCutOperation)
        exturdDef.SetDistanceExtent(10, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection)
        def.Features.ExtrudeFeatures.Add(exturdDef)

    End Sub
End Class

Public Class Selector

    Private WithEvents _interactEvents As InteractionEvents
    Private WithEvents _selectEvents As SelectEvents
    Private _stillSelecting As Boolean
    Private _inventor As Inventor.Application

    Public Sub New(ThisApplication As Inventor.Application)
        _inventor = ThisApplication
    End Sub

    Public Sub Pick()
        _stillSelecting = True

        _interactEvents = _inventor.CommandManager.CreateInteractionEvents
        _interactEvents.InteractionDisabled = False
        _interactEvents.StatusBarText = "Select point on face."
        _interactEvents.SetCursor(CursorTypeEnum.kCursorBuiltInCrosshair)

        _selectEvents = _interactEvents.SelectEvents
        _selectEvents.WindowSelectEnabled = False

        _interactEvents.Start()
        Do While _stillSelecting
            _inventor.UserInterfaceManager.DoEvents()
        Loop
        _interactEvents.Stop()

        _inventor.CommandManager.StopActiveCommand()

    End Sub

    Public Property SelectedFace As Object = Nothing
    Public Property ModelPosition As Point = Nothing

    Private Sub oSelectEvents_OnPreSelect(
            ByRef PreSelectEntity As Object,
            ByRef DoHighlight As Boolean,
            ByRef MorePreSelectEntities As ObjectCollection,
            SelectionDevice As SelectionDeviceEnum,
            ModelPosition As Point,
            ViewPosition As Point2d,
            View As Inventor.View) Handles _selectEvents.OnPreSelect

        DoHighlight = False

        If (not TypeOf PreSelectEntity Is Face) Then Return

        Dim face As Face = PreSelectEntity

        If (face.SurfaceType <> SurfaceTypeEnum.kPlaneSurface) Then Return

        DoHighlight = True

    End Sub

    Private Sub oInteractEvents_OnTerminate() Handles _interactEvents.OnTerminate
        _stillSelecting = False
    End Sub

    Private Sub oSelectEvents_OnSelect(
            ByVal JustSelectedEntities As ObjectsEnumerator,
            ByVal SelectionDevice As SelectionDeviceEnum,
            ByVal ModelPosition As Point,
            ByVal ViewPosition As Point2d,
            ByVal View As Inventor.View) Handles _selectEvents.OnSelect

        SelectedFace = JustSelectedEntities.Item(1)
        Me.ModelPosition = ModelPosition


        _stillSelecting = False
    End Sub
End Class