' Title: How to Check if Model State Exists Before Setting It in iLogic?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/how-to-check-if-model-state-exists-before-setting-it-in-ilogic/td-p/13783575#messageview_0
' Category: advanced
' Scraped: 2025-10-09T08:53:22.386621

Function DoesModelStateExist() as boolean
Dim oDoc as partdocument = Thisapplication.activeDocument

Dim pComp as partcomponentDefinition = odoc.componentDefinition

Dim DocModelStates as modelstates = PartComp.Modelstates

For each iModelstate as modelstate in Docmodelstates
    If ImodelState.name.contains("Name") 
      Return true
    else if modelstate.Name.contains("Name2") 
      return true
    else
     Return false
    end if 
next
End function