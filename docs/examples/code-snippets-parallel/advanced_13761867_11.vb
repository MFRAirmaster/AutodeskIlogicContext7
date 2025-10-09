' Title: what is the &quot;correct&quot; way to write this line of code for a rule?
' URL: https://forums.autodesk.com/t5/inventor-programming-forum/what-is-the-quot-correct-quot-way-to-write-this-line-of-code-for/td-p/13761867
' Category: advanced
' Scraped: 2025-10-09T09:01:32.182582

Dim topAsm As AssemblyDocument = ThisDoc.Document
Dim subAsmOcc As ComponentOccurrence = topAsm.ComponentDefinition.Occurrences(1) 'First occurrence of the top assembly
Dim subAsmDoc As AssemblyDocument = subAsmOcc.Definition.Document

Dim subAsmComponent As ICadComponent = Autodesk.iLogic.Interfaces.StandardObjectFactory.Create(subAsmDoc).Component

Dim isActive As Boolean = subAsmComponent.IsActive("Array")
subAsmComponent.IsActive("Array") = Not isActive' switch the state of Array