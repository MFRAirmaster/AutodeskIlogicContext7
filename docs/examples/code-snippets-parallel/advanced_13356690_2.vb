' Title: Addins - Using the Inventor Themes Light / Dark
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/addins-using-the-inventor-themes-light-dark/td-p/13356690
' Category: advanced
' Scraped: 2025-10-09T08:55:35.076159

public YourClassName()
{
    // This call is required by the designer.
    InitializeComponent();

    string DarkTheme = "inventor dark theme";
    string LightTheme = "inventor light theme";

    string AdeskLib = "Autodesk.iLogic.ThemeSkins.CustomSkins.";

    SkinManager.EnableFormSkins();

    var valDark = new SkinBlobXmlCreator(DarkTheme, AdeskLib, typeof(CustomThemeSkins).Assembly, null as string);
    var valLight = new SkinBlobXmlCreator(LightTheme, AdeskLib, typeof(CustomThemeSkins).Assembly, null as string);

    SkinManager.Default.RegisterSkin(valDark);
    SkinManager.Default.RegisterSkin(valLight);

    Inventor.ThemeManager oThemeManager = AddinGlobal.InventorApp.ThemeManager;
    Inventor.Theme oTheme = oThemeManager.ActiveTheme;

    switch (oTheme.Name)
    {
        case "LightTheme":
            this.LookAndFeel.SetSkinStyle(LightTheme);
            break;
        case "DarkTheme":
            this.LookAndFeel.SetSkinStyle(DarkTheme);
            break;
    }

    LookAndFeel.UpdateStyleSettings();
}