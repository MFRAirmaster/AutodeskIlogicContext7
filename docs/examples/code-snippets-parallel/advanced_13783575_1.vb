' Title: How to Check if Model State Exists Before Setting It in iLogic?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-check-if-model-state-exists-before-setting-it-in-ilogic/td-p/13783575#messageview_0
' Category: advanced
' Scraped: 2025-10-07T14:11:36.283135

If ModelStateExists(ModStateName & "'") Then    ActiveSheet.View("VIEW15").View.Suppressed = False
    ActiveSheet.View("VIEW15").NativeEntity.SetActiveModelState(ModStateName & "'")
Else
 ActiveSheet.View("VIEW15").View.Suppressed = TrueEnd If