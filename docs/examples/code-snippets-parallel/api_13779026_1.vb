' Title: ActiveView.Fit does not update until after the rule finishes running.
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/activeview-fit-does-not-update-until-after-the-rule-finishes/td-p/13779026
' Category: api
' Scraped: 2025-10-07T14:06:44.789693

Dim invApp As Inventor.Application = ThisApplication
invApp.ActiveView.Fit
invApp.ActiveView.SaveAsBitmap(savePath, 1150, 635)