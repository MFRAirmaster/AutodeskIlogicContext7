' Title: Promote a model state to become the master state with iLogic
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/promote-a-model-state-to-become-the-master-state-with-ilogic/td-p/12208768
' Category: advanced
' Scraped: 2025-10-07T13:05:48.627090

Sub Main 

Dim oDoc As PartDocument

'Dim oDoc As AssemblyDocement
oDoc = ThisApplication.ActiveDocument
    
Dim oDef As PartComponentDefinition
oDef = oDoc.ComponentDefinition

'Dim oDef As AssemblyComponentDefinition
'oDef = oDoc.ComponentDefinition
    
Dim oStates As ModelStates 	
oStates = oDef.ModelStates

Dim oTable As ModelStateTable	
oTable = oStates.ModelStateTable

'get the current model state, so we can return to it when done
oCurrentState = ThisDoc.ActiveModelState

MessageBox.Show(oCurrentState, "Title")



' Iterate All Rows of model states
i = 2
Dim oRow As ModelStateTableRow 

Dim oRowMaster As ModelStateTableRow
For Each oRow In oTable.TableRows
	'this activates the row to make it the current state
	
	If oCurrentState = oStates.Item(i).Name Then 
		
		oRowMaster =  oTable.TableRows.Item(i)
		
		Exit For
	End If 
	
'	MessageBox.Show( , "Title")
	
'	oStates.Item(i).Activate
	

	
	
	
	
	i=i+1
Next

oTable.TableRows.Item(1) = oRowMaster
	


	'set original state active again
	oStates.Item(oCurrentState).Activate 



iLogicVb.UpdateWhenDone = True
End Sub