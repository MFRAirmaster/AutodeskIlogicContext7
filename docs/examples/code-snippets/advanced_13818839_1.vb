' Title: iLogic - turn on/off parts
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/ilogic-turn-on-off-parts/td-p/13818839
' Category: advanced
' Scraped: 2025-10-07T13:27:31.016908

If Exterior_Option = "5/8" Then
	Feature.IsActive("1201-P Mainframe Head & Sill .625") = True
	Feature.IsActive("1201-P Mainframe jamb .625") = True
	Feature.IsActive("1201-P Mainframe Head & Sill 2") = False
	Feature.IsActive("1201-P Mainframe jamb 2") = False
	Feature.IsActive("1201-P Mainframe Head & Sill 3 1/2") = False
	Feature.IsActive("1201-P Mainframe jamb 3 1/2") = False