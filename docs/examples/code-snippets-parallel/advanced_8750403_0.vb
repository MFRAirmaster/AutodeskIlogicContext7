' Title: Create iProperties for all parts in Content Center
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/create-iproperties-for-all-parts-in-content-center/td-p/8750403#messageview_0
' Category: advanced
' Scraped: 2025-10-09T09:04:58.524984

trigger = iTrigger0
' Get the active assembly. 
    Dim oAsmDoc As AssemblyDocument
    oAsmDoc = ThisApplication.ActiveDocument

    ' Get the assembly component definition.
    Dim oAsmDef As AssemblyComponentDefinition
    oAsmDef = oAsmDoc.ComponentDefinition

    ' Get all of the leaf occurrences of the assembly.
    Dim oLeafOccs As ComponentOccurrencesEnumerator
    oLeafOccs = oAsmDef.Occurrences.AllLeafOccurrences

    ' Iterate through the occurrences and print the name.
    Dim oOcc As ComponentOccurrence
    For Each oOcc In oLeafOccs
     MsgBox(oOcc.ReferencedDocumentDescriptor.FullDocumentName)
     iProperties.Value(oOcc.Name,"Custom", "MojeHmotnost") = Round(oOcc.MassProperties.Mass,2) & " [kg]"
    Next