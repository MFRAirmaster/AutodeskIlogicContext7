' Title: Addins - Using the Inventor Themes Light / Dark
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/addins-using-the-inventor-themes-light-dark/td-p/13356690#messageview_0
' Category: advanced
' Scraped: 2025-10-07T13:32:57.851401

Public Sub New()

    ' This call is required by the designer.
    InitializeComponent()

    Dim DarkTheme As String = "inventor dark theme"
    Dim LightTheme As String = "inventor light theme"

    Dim AdeskLib As String = "Autodesk.iLogic.ThemeSkins.CustomSkins."

    SkinManager.EnableFormSkins()
    Dim valDark As New SkinBlobXmlCreator(DarkTheme, AdeskLib, GetType(CustomThemeSkins).Assembly, DirectCast(Nothing, String))
    Dim valLight As New SkinBlobXmlCreator(LightTheme, AdeskLib, GetType(CustomThemeSkins).Assembly, DirectCast(Nothing, String))

    SkinManager.Default.RegisterSkin(valDark)
    SkinManager.Default.RegisterSkin(valLight)

    Dim oThemeManager As Inventor.ThemeManager = AddinGlobal.InventorApp.ThemeManager
    Dim oTheme As Inventor.Theme = oThemeManager.ActiveTheme

    Select Case oTheme.Name
        Case "LightTheme"
            Me.LookAndFeel.SetSkinStyle(LightTheme)
        Case "DarkTheme"
            Me.LookAndFeel.SetSkinStyle(DarkTheme)
    End Select
    LookAndFeel.UpdateStyleSettings()

End Sub